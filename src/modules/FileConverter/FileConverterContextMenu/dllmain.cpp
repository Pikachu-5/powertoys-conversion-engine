#include "pch.h"

#include <nlohmann/json.hpp>

#include <algorithm>
#include <filesystem>
#include <string>
#include <vector>
#include <ShlObj.h>

using namespace Microsoft::WRL;

namespace
{
    constexpr wchar_t PIPE_NAME_PREFIX[] = L"\\\\.\\pipe\\powertoys_fileconverter_";
    constexpr char ACTION_FORMAT_CONVERT[] = "FormatConvert";
    constexpr char DESTINATION_PNG[] = "png";

    DWORD get_desktop_session_id()
    {
        DWORD session_id = 0;
        if (!ProcessIdToSessionId(GetCurrentProcessId(), &session_id))
        {
            session_id = WTSGetActiveConsoleSessionId();
        }

        return session_id;
    }

    std::wstring get_pipe_name()
    {
        return std::wstring(PIPE_NAME_PREFIX) + std::to_wstring(get_desktop_session_id());
    }

    std::string wide_to_utf8(const std::wstring& value)
    {
        if (value.empty())
        {
            return {};
        }

        const int required_size = WideCharToMultiByte(CP_UTF8, 0, value.c_str(), static_cast<int>(value.size()), nullptr, 0, nullptr, nullptr);
        if (required_size <= 0)
        {
            return {};
        }

        std::string utf8(required_size, '\0');
        const int written = WideCharToMultiByte(CP_UTF8, 0, value.c_str(), static_cast<int>(value.size()), utf8.data(), required_size, nullptr, nullptr);
        if (written <= 0)
        {
            return {};
        }

        return utf8;
    }

    bool is_existing_file_path(const std::wstring& path)
    {
        if (path.empty())
        {
            return false;
        }

        const DWORD attributes = GetFileAttributesW(path.c_str());
        return attributes != INVALID_FILE_ATTRIBUTES && (attributes & FILE_ATTRIBUTE_DIRECTORY) == 0;
    }

    bool is_supported_source_path(const std::wstring& path)
    {
        const wchar_t* extension = PathFindExtension(path.c_str());
        if (extension == nullptr || extension[0] == L'\0')
        {
            return false;
        }

        if (_wcsicmp(extension, L".png") == 0)
        {
            return false;
        }

#pragma warning(suppress : 26812)
        PERCEIVED perceived_type = PERCEIVED_TYPE_UNSPECIFIED;
        PERCEIVEDFLAG perceived_flags = PERCEIVEDFLAG_UNDEFINED;
        AssocGetPerceivedType(extension, &perceived_type, &perceived_flags, nullptr);
        return perceived_type == PERCEIVED_TYPE_IMAGE;
    }

    HRESULT send_pipe_request(const std::vector<std::wstring>& files)
    {
        nlohmann::json payload;
        payload["action"] = ACTION_FORMAT_CONVERT;
        payload["destination"] = DESTINATION_PNG;
        payload["files"] = nlohmann::json::array();

        for (const auto& file : files)
        {
            if (!is_existing_file_path(file))
            {
                continue;
            }

            const auto utf8_path = wide_to_utf8(file);
            if (!utf8_path.empty())
            {
                payload["files"].push_back(utf8_path);
            }
        }

        if (payload["files"].empty())
        {
            return S_OK;
        }

        const std::string body = payload.dump();
        const std::wstring pipe_name = get_pipe_name();

        if (!WaitNamedPipeW(pipe_name.c_str(), 2000))
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }

        HANDLE pipe_handle = CreateFileW(
            pipe_name.c_str(),
            GENERIC_WRITE,
            0,
            nullptr,
            OPEN_EXISTING,
            FILE_ATTRIBUTE_NORMAL,
            nullptr);

        if (pipe_handle == INVALID_HANDLE_VALUE)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }

        DWORD written = 0;
        const BOOL write_ok = WriteFile(pipe_handle, body.data(), static_cast<DWORD>(body.size()), &written, nullptr);
        const DWORD write_error = write_ok ? ERROR_SUCCESS : GetLastError();

        CloseHandle(pipe_handle);

        if (!write_ok)
        {
            return HRESULT_FROM_WIN32(write_error);
        }

        if (written != body.size())
        {
            return HRESULT_FROM_WIN32(ERROR_WRITE_FAULT);
        }

        return S_OK;
    }
}

HINSTANCE g_module_instance = 0;

BOOL APIENTRY DllMain(HMODULE module_handle, DWORD reason, LPVOID)
{
    if (reason == DLL_PROCESS_ATTACH)
    {
        g_module_instance = module_handle;
    }

    return TRUE;
}

class __declspec(uuid("57EC18F5-24D5-4DC6-AE2E-9D0F7A39F8BA")) FileConverterContextMenuCommand final :
    public RuntimeClass<RuntimeClassFlags<ClassicCom>, IExplorerCommand, IObjectWithSite, IShellExtInit, IContextMenu>
{
public:
    IFACEMETHODIMP GetTitle(_In_opt_ IShellItemArray*, _Outptr_result_nullonfailure_ PWSTR* name)
    {
        return SHStrDup(L"Convert to PNG", name);
    }

    IFACEMETHODIMP GetIcon(_In_opt_ IShellItemArray*, _Outptr_result_nullonfailure_ PWSTR* icon)
    {
        *icon = nullptr;
        return E_NOTIMPL;
    }

    IFACEMETHODIMP GetToolTip(_In_opt_ IShellItemArray*, _Outptr_result_nullonfailure_ PWSTR* info_tip)
    {
        *info_tip = nullptr;
        return E_NOTIMPL;
    }

    IFACEMETHODIMP GetCanonicalName(_Out_ GUID* guid_command_name)
    {
        *guid_command_name = __uuidof(this);
        return S_OK;
    }

    IFACEMETHODIMP GetState(_In_opt_ IShellItemArray* selection, _In_ BOOL, _Out_ EXPCMDSTATE* cmd_state)
    {
        *cmd_state = ECS_HIDDEN;

        if (selection == nullptr)
        {
            return S_OK;
        }

        std::vector<std::wstring> selected_paths;
        const HRESULT selected_paths_hr = get_selected_paths(selection, selected_paths);
        if (FAILED(selected_paths_hr))
        {
            return S_OK;
        }

        const bool has_supported_selection = std::any_of(
            selected_paths.begin(),
            selected_paths.end(),
            [](const std::wstring& path) {
                return is_supported_source_path(path);
            });

        if (has_supported_selection)
        {
            *cmd_state = ECS_ENABLED;
        }

        return S_OK;
    }

    IFACEMETHODIMP Invoke(_In_opt_ IShellItemArray* selection, _In_opt_ IBindCtx*)
    {
        if (selection == nullptr)
        {
            return S_OK;
        }

        std::vector<std::wstring> selected_paths;
        const HRESULT selected_paths_hr = get_selected_paths(selection, selected_paths);
        if (FAILED(selected_paths_hr))
        {
            return S_OK;
        }

        std::vector<std::wstring> supported_paths;
        for (const auto& path : selected_paths)
        {
            if (is_supported_source_path(path))
            {
                supported_paths.push_back(path);
            }
        }

        (void)send_pipe_request(supported_paths);
        return S_OK;
    }

    IFACEMETHODIMP GetFlags(_Out_ EXPCMDFLAGS* flags)
    {
        *flags = ECF_DEFAULT;
        return S_OK;
    }

    IFACEMETHODIMP EnumSubCommands(_COM_Outptr_ IEnumExplorerCommand** enum_commands)
    {
        *enum_commands = nullptr;
        return E_NOTIMPL;
    }

    IFACEMETHODIMP SetSite(_In_ IUnknown* site)
    {
        m_site = site;
        return S_OK;
    }

    IFACEMETHODIMP GetSite(_In_ REFIID riid, _COM_Outptr_ void** site)
    {
        return m_site.CopyTo(riid, site);
    }

    // IShellExtInit
    IFACEMETHODIMP Initialize(_In_opt_ PCIDLIST_ABSOLUTE, _In_opt_ IDataObject* data_object, _In_opt_ HKEY)
    {
        m_data_object = data_object;
        return S_OK;
    }

    // IContextMenu
    IFACEMETHODIMP QueryContextMenu(HMENU menu, UINT index_menu, UINT id_cmd_first, UINT, UINT flags)
    {
        if (menu == nullptr)
        {
            return E_INVALIDARG;
        }

        if ((flags & CMF_DEFAULTONLY) != 0 || m_data_object == nullptr)
        {
            return MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_NULL, 0);
        }

        std::vector<std::wstring> selected_paths;
        if (FAILED(get_selected_paths(m_data_object.Get(), selected_paths)))
        {
            return MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_NULL, 0);
        }

        const bool has_supported_selection = std::any_of(
            selected_paths.begin(),
            selected_paths.end(),
            [](const std::wstring& path) {
                return is_supported_source_path(path);
            });

        if (!has_supported_selection)
        {
            return MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_NULL, 0);
        }

        if (!InsertMenuW(menu, index_menu, MF_BYPOSITION | MF_STRING, id_cmd_first, L"Convert to PNG"))
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }

        return MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_NULL, 1);
    }

    IFACEMETHODIMP InvokeCommand(CMINVOKECOMMANDINFO* invoke_info)
    {
        if (invoke_info == nullptr || m_data_object == nullptr)
        {
            return S_OK;
        }

        if (!IS_INTRESOURCE(invoke_info->lpVerb) || LOWORD(invoke_info->lpVerb) != 0)
        {
            return S_OK;
        }

        std::vector<std::wstring> selected_paths;
        const HRESULT selected_paths_hr = get_selected_paths(m_data_object.Get(), selected_paths);
        if (FAILED(selected_paths_hr))
        {
            return S_OK;
        }

        std::vector<std::wstring> supported_paths;
        for (const auto& path : selected_paths)
        {
            if (is_supported_source_path(path))
            {
                supported_paths.push_back(path);
            }
        }

        (void)send_pipe_request(supported_paths);
        return S_OK;
    }

    IFACEMETHODIMP GetCommandString(UINT_PTR, UINT, UINT*, LPSTR, UINT)
    {
        return E_NOTIMPL;
    }

private:
    HRESULT get_selected_paths(IShellItemArray* selection, std::vector<std::wstring>& paths)
    {
        if (selection == nullptr)
        {
            return E_INVALIDARG;
        }

        DWORD item_count = 0;
        const HRESULT count_hr = selection->GetCount(&item_count);
        if (FAILED(count_hr))
        {
            return count_hr;
        }

        for (DWORD index = 0; index < item_count; ++index)
        {
            ComPtr<IShellItem> item;
            const HRESULT item_hr = selection->GetItemAt(index, &item);
            if (FAILED(item_hr))
            {
                continue;
            }

            PWSTR path_value = nullptr;
            const HRESULT path_hr = item->GetDisplayName(SIGDN_FILESYSPATH, &path_value);
            if (FAILED(path_hr))
            {
                continue;
            }

            if (path_value != nullptr && path_value[0] != L'\0')
            {
                paths.emplace_back(path_value);
            }

            if (path_value != nullptr)
            {
                CoTaskMemFree(path_value);
            }
        }

        return S_OK;
    }

    HRESULT get_selected_paths(IDataObject* data_object, std::vector<std::wstring>& paths)
    {
        if (data_object == nullptr)
        {
            return E_INVALIDARG;
        }

        ComPtr<IShellItemArray> shell_item_array;
        const HRESULT hr = SHCreateShellItemArrayFromDataObject(data_object, IID_PPV_ARGS(&shell_item_array));
        if (FAILED(hr))
        {
            return hr;
        }

        return get_selected_paths(shell_item_array.Get(), paths);
    }

    ComPtr<IUnknown> m_site;
    ComPtr<IDataObject> m_data_object;
};

CoCreatableClass(FileConverterContextMenuCommand)
CoCreatableClassWrlCreatorMapInclude(FileConverterContextMenuCommand)

STDAPI DllGetActivationFactory(_In_ HSTRING activatableClassId, _COM_Outptr_ IActivationFactory** factory)
{
    return Module<ModuleType::InProc>::GetModule().GetActivationFactory(activatableClassId, factory);
}

STDAPI DllCanUnloadNow()
{
    return Module<InProc>::GetModule().GetObjectCount() == 0 ? S_OK : S_FALSE;
}

STDAPI DllGetClassObject(_In_ REFCLSID rclsid, _In_ REFIID riid, _COM_Outptr_ void** instance)
{
    return Module<InProc>::GetModule().GetClassObject(rclsid, riid, instance);
}
