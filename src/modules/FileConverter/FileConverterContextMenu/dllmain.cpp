#include "pch.h"

#include <FileConversionEngine.h>

#include <filesystem>
#include <string>

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
    public RuntimeClass<RuntimeClassFlags<ClassicCom>, IExplorerCommand, IObjectWithSite>
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
            return path_hr;
        }

        const wchar_t* extension = PathFindExtension(path.c_str());
        if (extension == nullptr || extension[0] == L'\0')
        {
            return S_OK;
        }

#pragma warning(suppress : 26812)
        PERCEIVED type = PERCEIVED_TYPE_UNSPECIFIED;
        PERCEIVEDFLAG flags = PERCEIVEDFLAG_UNDEFINED;
        AssocGetPerceivedType(extension, &type, &flags, nullptr);
        if (type == PERCEIVED_TYPE_IMAGE)
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

        const std::filesystem::path input_path(path);
        std::filesystem::path output_path = input_path.parent_path() / input_path.stem();
        output_path += L"_converted.png";

        const auto conversion = file_converter::ConvertImageFile(input_path.wstring(), output_path.wstring(), file_converter::ImageFormat::Png);
        return conversion.hr;
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

private:
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

    ComPtr<IUnknown> m_site;
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
