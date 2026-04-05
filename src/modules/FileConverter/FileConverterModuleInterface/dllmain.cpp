#include "pch.h"

#include <FileConversionEngine.h>
#include <winrt/Windows.Foundation.Collections.h>
#include <common/utils/package.h>
#include <common/utils/process_path.h>
#include <common/utils/shell_ext_registration.h>
#include <interface/powertoy_module_interface.h>

#include <algorithm>
#include <filesystem>
#include <optional>
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
    constexpr wchar_t CONTEXT_MENU_PACKAGE_FILE_PREFIX[] = L"FileConverterContextMenuPackage";
    constexpr wchar_t CONTEXT_MENU_HANDLER_CLSID[] = L"{57EC18F5-24D5-4DC6-AE2E-9D0F7A39F8BA}";

    runtime_shell_ext::Spec build_win10_context_menu_spec()
    {
        runtime_shell_ext::Spec spec;
        spec.clsid = CONTEXT_MENU_HANDLER_CLSID;
        spec.sentinelKey = L"Software\\Microsoft\\PowerToys\\FileConverter";
        spec.sentinelValue = L"ContextMenuRegisteredWin10";
        spec.dllFileCandidates = {
            L"WinUI3Apps\\PowerToys.FileConverterContextMenu.dll",
            L"PowerToys.FileConverterContextMenu.dll",
        };
        spec.friendlyName = L"File Converter Context Menu";
        spec.systemFileAssocHandlerName = L"FileConverterContextMenu";
        spec.representativeSystemExt = L".bmp";
        spec.systemFileAssocExtensions = {
            L".bmp",
            L".dib",
            L".gif",
            L".jfif",
            L".jpe",
            L".jpeg",
            L".jpg",
            L".jxr",
            L".tif",
            L".tiff",
            L".wdp",
            L".heic",
            L".heif",
            L".webp",
        };
        return spec;
    }

    std::optional<std::filesystem::path> find_latest_context_menu_package(const std::filesystem::path& context_menu_path)
    {
        const std::filesystem::path stable_package_path = context_menu_path / CONTEXT_MENU_PACKAGE_FILE_NAME;
        if (std::filesystem::exists(stable_package_path))
        {
            return stable_package_path;
        }

        std::vector<std::filesystem::path> candidate_packages;
        std::error_code ec;
        for (std::filesystem::directory_iterator it(context_menu_path, ec); !ec && it != std::filesystem::directory_iterator(); it.increment(ec))
        {
            if (!it->is_regular_file(ec))
            {
                continue;
            }

            const auto file_name = it->path().filename().wstring();
            const auto extension = it->path().extension().wstring();
            if (_wcsicmp(extension.c_str(), L".msix") != 0)
            {
                continue;
            }

            if (file_name.rfind(CONTEXT_MENU_PACKAGE_FILE_PREFIX, 0) == 0)
            {
                candidate_packages.push_back(it->path());
            }
        }

        if (candidate_packages.empty())
        {
            return std::nullopt;
        }

        std::sort(candidate_packages.begin(), candidate_packages.end(), [](const auto& lhs, const auto& rhs) {
            std::error_code lhs_ec;
            std::error_code rhs_ec;
            const auto lhs_time = std::filesystem::last_write_time(lhs, lhs_ec);
            const auto rhs_time = std::filesystem::last_write_time(rhs, rhs_ec);

            if (lhs_ec && rhs_ec)
            {
                return lhs.wstring() < rhs.wstring();
            }

            if (lhs_ec)
            {
                return true;
            }

            if (rhs_ec)
            {
                return false;
            }

            return lhs_time < rhs_time;
        });

        return candidate_packages.back();
    }

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

        const std::filesystem::path module_path = get_module_folderpath(reinterpret_cast<HMODULE>(&__ImageBase));
        const std::filesystem::path context_menu_path = module_path / L"WinUI3Apps";
        if (!std::filesystem::exists(context_menu_path))
        {
            return;
        }

        const auto package_path = find_latest_context_menu_package(context_menu_path);
        if (!package_path.has_value())
        {
            return;
        }

        if (!package::IsPackageRegisteredWithPowerToysVersion(CONTEXT_MENU_PACKAGE_DISPLAY_NAME))
        {
            (void)package::RegisterSparsePackage(context_menu_path.wstring(), package_path->wstring());
        }
    }

    void ensure_context_menu_runtime_registered()
    {
        if (package::IsWin11OrGreater())
        {
            return;
        }

        (void)runtime_shell_ext::EnsureRegistered(build_win10_context_menu_spec(), reinterpret_cast<HMODULE>(&__ImageBase));
    }

    void unregister_context_menu_runtime()
    {
        if (package::IsWin11OrGreater())
        {
            return;
        }

        runtime_shell_ext::Unregister(build_win10_context_menu_spec());
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
        ensure_context_menu_runtime_registered();
        m_enabled = true;
    }

    void disable() override
    {
        unregister_context_menu_runtime();
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
