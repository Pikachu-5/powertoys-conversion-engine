#include "pch.h"

#include <ShlObj.h>
#include <winrt/Windows.Data.Json.h>
#include <winrt/Windows.Foundation.Collections.h>
#include <winrt/base.h>

#include <string>
#include <vector>

using namespace Microsoft::WRL;
namespace json = winrt::Windows::Data::Json;

namespace
{
    constexpr wchar_t PIPE_NAME_PREFIX[] = L"\\\\.\\pipe\\powertoys_fileconverter_";
    constexpr DWORD PIPE_CONNECT_TIMEOUT_MS = 1000;

    std::wstring get_pipe_name_for_current_session()
    {
        DWORD session_id = 0;
        if (!ProcessIdToSessionId(GetCurrentProcessId(), &session_id))
        {
            session_id = 0;
        }

        return std::wstring(PIPE_NAME_PREFIX) + std::to_wstring(session_id);
    }

    HRESULT get_selected_paths(IShellItemArray* selection, std::vector<std::wstring>& paths)
    {
        if (selection == nullptr)
        {
            return E_INVALIDARG;
        }

        paths.clear();

        DWORD count = 0;
        const HRESULT count_hr = selection->GetCount(&count);
        if (FAILED(count_hr))
        {
            return count_hr;
        }

        for (DWORD i = 0; i < count; ++i)
        {
            ComPtr<IShellItem> item;
            const HRESULT item_hr = selection->GetItemAt(i, &item);
            if (FAILED(item_hr) || item == nullptr)
            {
                continue;
            }

            PWSTR path_value = nullptr;
            const HRESULT path_hr = item->GetDisplayName(SIGDN_FILESYSPATH, &path_value);
            if (FAILED(path_hr) || path_value == nullptr || path_value[0] == L'\0')
            {
                if (path_value != nullptr)
                {
                    CoTaskMemFree(path_value);
                }

                continue;
            }

            paths.emplace_back(path_value);
            CoTaskMemFree(path_value);
        }

        return paths.empty() ? E_FAIL : S_OK;
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

    bool should_enable_for_path(const std::wstring& path)
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

    bool should_enable_for_paths(const std::vector<std::wstring>& paths)
    {
        if (paths.empty())
        {
            return false;
        }

        for (const auto& path : paths)
        {
            if (!should_enable_for_path(path))
            {
                return false;
            }
        }

        return true;
    }

    std::string build_format_convert_payload(const std::vector<std::wstring>& paths)
    {
        json::JsonObject payload;
        payload.Insert(L"action", json::JsonValue::CreateStringValue(L"FormatConvert"));
        payload.Insert(L"destination", json::JsonValue::CreateStringValue(L"png"));

        json::JsonArray files;
        for (const auto& path : paths)
        {
            files.Append(json::JsonValue::CreateStringValue(path));
        }

        payload.Insert(L"files", files);
        return winrt::to_string(payload.Stringify());
    }

    HRESULT send_format_convert_request(const std::vector<std::wstring>& paths)
    {
        if (!should_enable_for_paths(paths))
        {
            return E_INVALIDARG;
        }

        const std::wstring pipe_name = get_pipe_name_for_current_session();
        if (!WaitNamedPipeW(pipe_name.c_str(), PIPE_CONNECT_TIMEOUT_MS))
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }

        HANDLE pipe_handle = CreateFileW(
            pipe_name.c_str(),
            GENERIC_WRITE,
            0,
            nullptr,
            OPEN_EXISTING,
            0,
            nullptr);

        if (pipe_handle == INVALID_HANDLE_VALUE)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }

        const std::string payload = build_format_convert_payload(paths);

        DWORD bytes_written = 0;
        const BOOL write_result = WriteFile(
            pipe_handle,
            payload.data(),
            static_cast<DWORD>(payload.size()),
            &bytes_written,
            nullptr);

        const DWORD write_error = write_result ? ERROR_SUCCESS : GetLastError();
        CloseHandle(pipe_handle);

        if (!write_result || bytes_written != payload.size())
        {
            return HRESULT_FROM_WIN32(write_result ? ERROR_WRITE_FAULT : write_error);
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

        std::vector<std::wstring> paths;
        if (FAILED(get_selected_paths(selection, paths)))
        {
            return S_OK;
        }

        if (should_enable_for_paths(paths))
        {
            *cmd_state = ECS_ENABLED;
        }

        return S_OK;
    }

    IFACEMETHODIMP Invoke(_In_opt_ IShellItemArray* selection, _In_opt_ IBindCtx*)
    {
        if (selection != nullptr)
        {
            std::vector<std::wstring> paths;
            if (SUCCEEDED(get_selected_paths(selection, paths)))
            {
                (void)send_format_convert_request(paths);
            }
        }

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

    IFACEMETHODIMP Initialize(_In_opt_ PCIDLIST_ABSOLUTE, _In_opt_ IDataObject* data_object, _In_opt_ HKEY)
    {
        m_data_object = data_object;
        return S_OK;
    }

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

        std::vector<std::wstring> paths;
        if (FAILED(get_selected_paths(m_data_object.Get(), paths)) || !should_enable_for_paths(paths))
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

        std::vector<std::wstring> paths;
        if (FAILED(get_selected_paths(m_data_object.Get(), paths)) || !should_enable_for_paths(paths))
        {
            return S_OK;
        }

        (void)send_format_convert_request(paths);
        return S_OK;
    }

    IFACEMETHODIMP GetCommandString(UINT_PTR, UINT, UINT*, LPSTR, UINT)
    {
        return E_NOTIMPL;
    }

private:
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
