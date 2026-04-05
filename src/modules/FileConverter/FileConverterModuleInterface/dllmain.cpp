#include "pch.h"

#include <FileConversionEngine.h>
#include <winrt/Windows.Foundation.Collections.h>
#include <common/utils/package.h>
#include <common/utils/process_path.h>
#include <interface/powertoy_module_interface.h>

#include <filesystem>
#include <string>
#include <vector>

extern "C" IMAGE_DOS_HEADER __ImageBase;

namespace
{
    constexpr wchar_t MODULE_NAME[] = L"File Converter";
    constexpr wchar_t MODULE_KEY[] = L"FileConverter";
    constexpr wchar_t MODULE_CONFIG[] = LR"({"name":"File Converter","version":"1.0","properties":{}})";
    constexpr wchar_t CONTEXT_MENU_PACKAGE_DISPLAY_NAME[] = L"FileConverterContextMenu";
    constexpr wchar_t CONTEXT_MENU_PACKAGE_FILE_NAME[] = L"FileConverterContextMenuPackage.msix";

    std::vector<std::wstring> split(const std::wstring& source, wchar_t delimiter)
    {
        std::vector<std::wstring> result;
        std::wstring current;

        for (const wchar_t ch : source)
        {
            if (ch == delimiter)
            {
                result.push_back(current);
                current.clear();
                continue;
            }

            current.push_back(ch);
        }

        result.push_back(current);
        return result;
    }

    file_converter::ImageFormat parse_format(const std::wstring& value)
    {
        if (_wcsicmp(value.c_str(), L"jpeg") == 0 || _wcsicmp(value.c_str(), L"jpg") == 0)
        {
            return file_converter::ImageFormat::Jpeg;
        }

        if (_wcsicmp(value.c_str(), L"bmp") == 0)
        {
            return file_converter::ImageFormat::Bmp;
        }

        if (_wcsicmp(value.c_str(), L"tiff") == 0 || _wcsicmp(value.c_str(), L"tif") == 0)
        {
            return file_converter::ImageFormat::Tiff;
        }

        return file_converter::ImageFormat::Png;
    }

    void ensure_context_menu_package_registered()
    {
        if (!package::IsWin11OrGreater())
        {
            return;
        }

        const std::wstring module_path = get_module_folderpath(reinterpret_cast<HMODULE>(&__ImageBase));
        const std::wstring package_uri = module_path + L"\\" + CONTEXT_MENU_PACKAGE_FILE_NAME;
        if (!std::filesystem::exists(package_uri))
        {
            return;
        }

        if (!package::IsPackageRegisteredWithPowerToysVersion(CONTEXT_MENU_PACKAGE_DISPLAY_NAME))
        {
            (void)package::RegisterSparsePackage(module_path, package_uri);
        }
    }
}

class FileConverterModule : public PowertoyModuleIface
{
public:
    void destroy() override
    {
        delete this;
    }

    const wchar_t* get_name() override
    {
        return MODULE_NAME;
    }

    const wchar_t* get_key() override
    {
        return MODULE_KEY;
    }

    bool get_config(wchar_t* buffer, int* buffer_size) override
    {
        const int required_size = static_cast<int>(wcslen(MODULE_CONFIG) + 1);
        if (buffer == nullptr || *buffer_size < required_size)
        {
            *buffer_size = required_size;
            return false;
        }

        wcscpy_s(buffer, *buffer_size, MODULE_CONFIG);
        return true;
    }

    void set_config(const wchar_t* /*config*/) override
    {
    }

    void call_custom_action(const wchar_t* action) override
    {
        if (action == nullptr)
        {
            return;
        }

        // Temporary action format for Phase 2 scaffold:
        // convert:<input_path>|<output_path>|<format>
        std::wstring action_text = action;
        constexpr wchar_t action_prefix[] = L"convert:";
        if (action_text.rfind(action_prefix, 0) != 0)
        {
            return;
        }

        std::wstring payload = action_text.substr(wcslen(action_prefix));
        const auto parts = split(payload, L'|');
        if (parts.size() < 3)
        {
            return;
        }

        const auto format = parse_format(parts[2]);
        (void)file_converter::ConvertImageFile(parts[0], parts[1], format);
    }

    void enable() override
    {
        ensure_context_menu_package_registered();
        m_enabled = true;
    }

    void disable() override
    {
        m_enabled = false;
    }

    bool is_enabled() override
    {
        return m_enabled;
    }

private:
    bool m_enabled = false;
};

extern "C" __declspec(dllexport) PowertoyModuleIface* __cdecl powertoy_create()
{
    return new FileConverterModule();
}
