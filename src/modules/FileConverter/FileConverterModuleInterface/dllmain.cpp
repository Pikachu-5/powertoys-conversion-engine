#include "pch.h"

#include <FileConversionEngine.h>
#include <winrt/Windows.Data.Json.h>
#include <winrt/Windows.Foundation.Collections.h>
#include <winrt/base.h>
#include <common/logger/logger.h>
#include <common/utils/package.h>
#include <common/utils/process_path.h>
#include <common/utils/shell_ext_registration.h>
#include <interface/powertoy_module_interface.h>

#include <algorithm>
#include <atomic>
#include <chrono>
#include <condition_variable>
#include <cwctype>
#include <filesystem>
#include <mutex>
#include <optional>
#include <queue>
#include <string>
#include <thread>
#include <unordered_set>
#include <vector>

extern "C" IMAGE_DOS_HEADER __ImageBase;
namespace json = winrt::Windows::Data::Json;

namespace
{
    constexpr wchar_t MODULE_NAME[] = L"File Converter";
    constexpr wchar_t MODULE_KEY[] = L"FileConverter";
    constexpr wchar_t MODULE_CONFIG[] = LR"({"name":"File Converter","version":"1.0","properties":{}})";
    constexpr wchar_t CONTEXT_MENU_PACKAGE_DISPLAY_NAME[] = L"FileConverterContextMenu";
    constexpr wchar_t CONTEXT_MENU_PACKAGE_FILE_NAME[] = L"FileConverterContextMenuPackage.msix";
    constexpr wchar_t CONTEXT_MENU_PACKAGE_FILE_PREFIX[] = L"FileConverterContextMenuPackage";
    constexpr wchar_t CONTEXT_MENU_HANDLER_CLSID[] = L"{57EC18F5-24D5-4DC6-AE2E-9D0F7A39F8BA}";
    constexpr wchar_t PIPE_NAME_PREFIX[] = L"\\\\.\\pipe\\powertoys_fileconverter_";
    constexpr wchar_t ACTION_FORMAT_CONVERT[] = L"FormatConvert";

    struct ConversionRequest
    {
        file_converter::ImageFormat format = file_converter::ImageFormat::Png;
        std::vector<std::wstring> files;
        size_t skipped_entries = 0;
    };

    struct ConversionSummary
    {
        size_t succeeded = 0;
        size_t missing_inputs = 0;
        size_t failed = 0;
        std::wstring first_failed_path;
        std::wstring first_failed_error;
    };

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

    std::wstring to_lower(std::wstring value)
    {
        std::transform(value.begin(), value.end(), value.begin(), [](wchar_t ch) {
            return static_cast<wchar_t>(towlower(ch));
        });

        return value;
    }

    file_converter::ImageFormat parse_format(const std::wstring& value)
    {
        const std::wstring lower = to_lower(value);

        if (lower == L"jpeg" || lower == L"jpg")
        {
            return file_converter::ImageFormat::Jpeg;
        }

        if (lower == L"bmp")
        {
            return file_converter::ImageFormat::Bmp;
        }

        if (lower == L"tiff" || lower == L"tif")
        {
            return file_converter::ImageFormat::Tiff;
        }

        return file_converter::ImageFormat::Png;
    }

    std::wstring extension_for_format(file_converter::ImageFormat format)
    {
        switch (format)
        {
        case file_converter::ImageFormat::Jpeg:
            return L".jpg";
        case file_converter::ImageFormat::Bmp:
            return L".bmp";
        case file_converter::ImageFormat::Tiff:
            return L".tiff";
        case file_converter::ImageFormat::Png:
        default:
            return L".png";
        }
    }

    std::wstring get_pipe_name_for_current_session()
    {
        DWORD session_id = 0;
        if (!ProcessIdToSessionId(GetCurrentProcessId(), &session_id))
        {
            session_id = 0;
        }

        return std::wstring(PIPE_NAME_PREFIX) + std::to_wstring(session_id);
    }

    std::string read_pipe_message(HANDLE pipe_handle)
    {
        constexpr DWORD BUFFER_SIZE = 4096;
        char buffer[BUFFER_SIZE] = {};
        std::string payload;

        while (true)
        {
            DWORD bytes_read = 0;
            const BOOL read_ok = ReadFile(pipe_handle, buffer, BUFFER_SIZE, &bytes_read, nullptr);
            if (bytes_read > 0)
            {
                payload.append(buffer, bytes_read);
            }

            if (read_ok)
            {
                break;
            }

            const DWORD read_error = GetLastError();
            if (read_error == ERROR_MORE_DATA)
            {
                continue;
            }

            if (read_error == ERROR_BROKEN_PIPE || read_error == ERROR_PIPE_NOT_CONNECTED)
            {
                break;
            }

            Logger::warn(L"File Converter pipe read failed. Error={}", read_error);
            payload.clear();
            break;
        }

        return payload;
    }

    bool try_parse_format_convert_request(
        const std::string& payload,
        ConversionRequest& request,
        std::wstring& rejection_reason)
    {
        request = {};
        rejection_reason.clear();

        if (payload.empty())
        {
            rejection_reason = L"empty payload";
            return false;
        }

        json::JsonObject json_payload;
        if (!json::JsonObject::TryParse(winrt::to_hstring(payload), json_payload))
        {
            rejection_reason = L"invalid JSON";
            return false;
        }

        if (!json_payload.HasKey(L"action"))
        {
            rejection_reason = L"missing action";
            return false;
        }

        const auto action_value = json_payload.GetNamedValue(L"action");
        if (action_value.ValueType() != json::JsonValueType::String)
        {
            rejection_reason = L"action is not a string";
            return false;
        }

        const auto action = json_payload.GetNamedString(L"action");
        if (_wcsicmp(action.c_str(), ACTION_FORMAT_CONVERT) != 0)
        {
            rejection_reason = L"unsupported action";
            return false;
        }

        std::wstring destination = L"png";
        if (json_payload.HasKey(L"destination"))
        {
            const auto destination_value = json_payload.GetNamedValue(L"destination");
            if (destination_value.ValueType() == json::JsonValueType::String)
            {
                destination = json_payload.GetNamedString(L"destination").c_str();
            }
        }

        if (!json_payload.HasKey(L"files"))
        {
            rejection_reason = L"missing files array";
            return false;
        }

        const auto files_value = json_payload.GetNamedValue(L"files");
        if (files_value.ValueType() != json::JsonValueType::Array)
        {
            rejection_reason = L"files is not an array";
            return false;
        }

        const auto files_array = json_payload.GetNamedArray(L"files");
        for (const auto& file_value : files_array)
        {
            if (file_value.ValueType() != json::JsonValueType::String)
            {
                ++request.skipped_entries;
                continue;
            }

            const auto file_path = file_value.GetString();
            if (file_path.empty())
            {
                ++request.skipped_entries;
                continue;
            }

            request.files.push_back(file_path.c_str());
        }

        if (request.files.empty())
        {
            rejection_reason = L"no valid file paths";
            return false;
        }

        request.format = parse_format(destination);
        return true;
    }

    ConversionSummary process_format_convert_request(const ConversionRequest& request)
    {
        ConversionSummary summary;
        const std::wstring output_extension = extension_for_format(request.format);
        std::unordered_set<std::wstring> seen_files;

        for (const auto& file : request.files)
        {
            if (!seen_files.insert(file).second)
            {
                continue;
            }

            const std::filesystem::path input_path(file);
            std::error_code ec;
            if (input_path.empty() || !std::filesystem::exists(input_path, ec) || ec)
            {
                ++summary.missing_inputs;
                continue;
            }

            ec.clear();
            if (!std::filesystem::is_regular_file(input_path, ec) || ec)
            {
                ++summary.missing_inputs;
                continue;
            }

            std::filesystem::path output_path = input_path.parent_path() / input_path.stem();
            output_path += L"_converted";
            output_path += output_extension;

            const auto conversion = file_converter::ConvertImageFile(input_path.wstring(), output_path.wstring(), request.format);
            if (conversion.succeeded())
            {
                ++summary.succeeded;
                continue;
            }

            ++summary.failed;
            if (summary.first_failed_path.empty())
            {
                summary.first_failed_path = input_path.wstring();
                summary.first_failed_error = conversion.error_message;
            }
        }

        return summary;
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

    class FileConverterPipeOrchestrator
    {
    public:
        void start(const std::wstring& pipe_name)
        {
            if (m_running.exchange(true))
            {
                return;
            }

            m_pipe_name = pipe_name;
            m_listener_thread = std::thread(&FileConverterPipeOrchestrator::listener_loop, this);
            m_worker_thread = std::thread(&FileConverterPipeOrchestrator::worker_loop, this);
        }

        void stop()
        {
            if (!m_running.exchange(false))
            {
                return;
            }

            wake_listener();
            m_queue_cv.notify_all();

            if (m_listener_thread.joinable())
            {
                m_listener_thread.join();
            }

            if (m_worker_thread.joinable())
            {
                m_worker_thread.join();
            }

            std::queue<std::string> empty;
            {
                std::scoped_lock lock(m_queue_mutex);
                std::swap(m_pending_payloads, empty);
            }
        }

        void enqueue_action_payload(std::string payload)
        {
            if (!m_running.load())
            {
                return;
            }

            enqueue_payload(std::move(payload));
        }

        ~FileConverterPipeOrchestrator()
        {
            stop();
        }

    private:
        void wake_listener() const
        {
            if (m_pipe_name.empty())
            {
                return;
            }

            HANDLE wake_handle = CreateFileW(
                m_pipe_name.c_str(),
                GENERIC_WRITE,
                0,
                nullptr,
                OPEN_EXISTING,
                0,
                nullptr);

            if (wake_handle != INVALID_HANDLE_VALUE)
            {
                CloseHandle(wake_handle);
            }
        }

        void enqueue_payload(std::string payload)
        {
            {
                std::scoped_lock lock(m_queue_mutex);
                m_pending_payloads.push(std::move(payload));
            }

            m_queue_cv.notify_one();
        }

        void process_payload(const std::string& payload)
        {
            ConversionRequest request;
            std::wstring rejection_reason;
            if (!try_parse_format_convert_request(payload, request, rejection_reason))
            {
                if (!rejection_reason.empty())
                {
                    Logger::warn(L"File Converter ignored malformed request: {}", rejection_reason);
                }

                return;
            }

            const auto summary = process_format_convert_request(request);

            if (request.skipped_entries > 0)
            {
                Logger::warn(L"File Converter request skipped {} invalid file entries.", request.skipped_entries);
            }

            if (summary.missing_inputs > 0)
            {
                Logger::warn(L"File Converter request skipped {} missing input files.", summary.missing_inputs);
            }

            if (summary.failed > 0)
            {
                Logger::warn(L"File Converter conversion failed for {} file(s).", summary.failed);
                if (!summary.first_failed_path.empty())
                {
                    Logger::warn(L"First conversion failure: path='{}' reason='{}'", summary.first_failed_path, summary.first_failed_error);
                }
            }
        }

        void worker_loop()
        {
            while (true)
            {
                std::string payload;
                {
                    std::unique_lock lock(m_queue_mutex);
                    m_queue_cv.wait(lock, [this] {
                        return !m_running.load() || !m_pending_payloads.empty();
                    });

                    if (m_pending_payloads.empty())
                    {
                        if (!m_running.load())
                        {
                            break;
                        }

                        continue;
                    }

                    payload = std::move(m_pending_payloads.front());
                    m_pending_payloads.pop();
                }

                process_payload(payload);
            }
        }

        void listener_loop()
        {
            while (m_running.load())
            {
                HANDLE pipe_handle = CreateNamedPipeW(
                    m_pipe_name.c_str(),
                    PIPE_ACCESS_INBOUND,
                    PIPE_TYPE_MESSAGE | PIPE_READMODE_MESSAGE | PIPE_WAIT,
                    PIPE_UNLIMITED_INSTANCES,
                    0,
                    4096,
                    0,
                    nullptr);

                if (pipe_handle == INVALID_HANDLE_VALUE)
                {
                    std::this_thread::sleep_for(std::chrono::milliseconds(100));
                    continue;
                }

                const BOOL connected = ConnectNamedPipe(pipe_handle, nullptr) ? TRUE : (GetLastError() == ERROR_PIPE_CONNECTED);
                if (!connected)
                {
                    CloseHandle(pipe_handle);
                    if (m_running.load())
                    {
                        std::this_thread::sleep_for(std::chrono::milliseconds(50));
                    }

                    continue;
                }

                const std::string payload = read_pipe_message(pipe_handle);

                FlushFileBuffers(pipe_handle);
                DisconnectNamedPipe(pipe_handle);
                CloseHandle(pipe_handle);

                if (!m_running.load())
                {
                    break;
                }

                if (!payload.empty())
                {
                    enqueue_payload(payload);
                }
            }
        }

        std::atomic<bool> m_running = false;
        std::wstring m_pipe_name;
        std::thread m_listener_thread;
        std::thread m_worker_thread;
        std::mutex m_queue_mutex;
        std::condition_variable m_queue_cv;
        std::queue<std::string> m_pending_payloads;
    };
}

class FileConverterModule : public PowertoyModuleIface
{
public:
    ~FileConverterModule()
    {
        disable();
    }

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

        m_pipe_orchestrator.enqueue_action_payload(winrt::to_string(action));
    }

    void enable() override
    {
        if (m_enabled)
        {
            return;
        }

        ensure_context_menu_package_registered();
        ensure_context_menu_runtime_registered();
        m_pipe_orchestrator.start(get_pipe_name_for_current_session());
        m_enabled = true;
    }

    void disable() override
    {
        if (!m_enabled)
        {
            return;
        }

        m_pipe_orchestrator.stop();
        unregister_context_menu_runtime();
        m_enabled = false;
    }

    bool is_enabled() override
    {
        return m_enabled;
    }

private:
    bool m_enabled = false;
    FileConverterPipeOrchestrator m_pipe_orchestrator;
};

extern "C" __declspec(dllexport) PowertoyModuleIface* __cdecl powertoy_create()
{
    return new FileConverterModule();
}
