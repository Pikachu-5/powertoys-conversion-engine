#include "pch.h"

#include <FileConversionEngine.h>

#include <filesystem>
#include <string>
#include <ShlObj.h>

using namespace Microsoft::WRL;

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

        std::wstring path;
        const HRESULT path_hr = get_first_selected_path(selection, path);
        if (FAILED(path_hr))
        {
            return S_OK;
        }

        if (should_enable_for_path(path))
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

        std::wstring path;
        const HRESULT path_hr = get_first_selected_path(selection, path);
        if (FAILED(path_hr))
        {
            return path_hr;
        }

        if (!should_enable_for_path(path))
        {
            return E_INVALIDARG;
        }

        return convert_to_png(path);
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

        std::wstring path;
        if (FAILED(get_first_selected_path(m_data_object.Get(), path)) || !should_enable_for_path(path))
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
            return E_INVALIDARG;
        }

        if (!IS_INTRESOURCE(invoke_info->lpVerb) || LOWORD(invoke_info->lpVerb) != 0)
        {
            return E_FAIL;
        }

        std::wstring path;
        const auto path_hr = get_first_selected_path(m_data_object.Get(), path);
        if (FAILED(path_hr))
        {
            return path_hr;
        }

        if (!should_enable_for_path(path))
        {
            return E_INVALIDARG;
        }

        return convert_to_png(path);
    }

    IFACEMETHODIMP GetCommandString(UINT_PTR, UINT, UINT*, LPSTR, UINT)
    {
        return E_NOTIMPL;
    }

private:
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
        PERCEIVED type = PERCEIVED_TYPE_UNSPECIFIED;
        PERCEIVEDFLAG flags = PERCEIVEDFLAG_UNDEFINED;
        AssocGetPerceivedType(extension, &type, &flags, nullptr);
        return type == PERCEIVED_TYPE_IMAGE;
    }

    HRESULT convert_to_png(const std::wstring& path)
    {
        const std::filesystem::path input_path(path);
        std::filesystem::path output_path = input_path.parent_path() / input_path.stem();
        output_path += L"_converted.png";

        const auto conversion = file_converter::ConvertImageFile(input_path.wstring(), output_path.wstring(), file_converter::ImageFormat::Png);
        return conversion.hr;
    }

    HRESULT get_first_selected_path(IShellItemArray* selection, std::wstring& path)
    {
        ComPtr<IShellItem> item;
        HRESULT hr = selection->GetItemAt(0, &item);
        if (FAILED(hr))
        {
            return hr;
        }

        PWSTR path_value = nullptr;
        hr = item->GetDisplayName(SIGDN_FILESYSPATH, &path_value);
        if (FAILED(hr))
        {
            return hr;
        }

        if (path_value == nullptr || path_value[0] == L'\0')
        {
            if (path_value != nullptr)
            {
                CoTaskMemFree(path_value);
            }

            return E_FAIL;
        }

        path.assign(path_value);
        CoTaskMemFree(path_value);

        return S_OK;
    }

    HRESULT get_first_selected_path(IDataObject* data_object, std::wstring& path)
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

        return get_first_selected_path(shell_item_array.Get(), path);
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
