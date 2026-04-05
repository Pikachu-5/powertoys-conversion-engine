#include "pch.h"

#include <FileConversionEngine.h>
<<<<<<< Updated upstream
#include <nlohmann/json.hpp>
#include <winrt/Windows.Foundation.Collections.h>
=======
#include <winrt/Windows.Data.Json.h>
#include <winrt/Windows.Foundation.Collections.h>
#include <winrt/base.h>
>>>>>>> Stashed changes
#include <common/logger/logger.h>
#include <common/utils/package.h>
#include <common/utils/process_path.h>
#include <common/utils/shell_ext_registration.h>
#include <interface/powertoy_module_interface.h>

#include <algorithm>
#include <atomic>
#include <chrono>
<<<<<<< Updated upstream
#include <condition_variable>
#include <cwctype>
=======
>>>>>>> Stashed changes
#include <filesystem>
#include <mutex>
#include <optional>
#include <queue>
#include <string>
#include <thread>
<<<<<<< Updated upstream
#include <unordered_set>
=======
>>>>>>> Stashed changes
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
<<<<<<< Updated upstream
    constexpr char ACTION_FORMAT_CONVERT[] = "FormatConvert";
    constexpr size_t PIPE_READ_BUFFER_SIZE = 8192;
=======

    struct ConversionSummary
    {
        size_t succeeded = 0;
        size_t missing_inputs = 0;
        size_t failed = 0;
        std::wstring first_failed_path;
        std::wstring first_failed_error;
    };
>>>>>>> Stashed changes

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

<<<<<<< Updated upstream
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

    std::wstring utf8_to_wide(const std::string& value)
    {
        if (value.empty())
        {
            return {};
        }

        const int required_size = MultiByteToWideChar(CP_UTF8, 0, value.data(), static_cast<int>(value.size()), nullptr, 0);
        if (required_size <= 0)
        {
            return {};
        }

        std::wstring wide(required_size, L'\0');
        const int converted = MultiByteToWideChar(CP_UTF8, 0, value.data(), static_cast<int>(value.size()), wide.data(), required_size);
        if (converted <= 0)
        {
            return {};
        }

        return wide;
    }

    std::wstring to_lower(std::wstring value)
    {
        std::transform(value.begin(), value.end(), value.begin(), [](wchar_t ch) {
            return static_cast<wchar_t>(towlower(ch));
        });

        return value;
    }

=======
>>>>>>> Stashed changes
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

<<<<<<< Updated upstream
    std::filesystem::path build_output_path(const std::filesystem::path& input_path, const std::wstring& destination)
    {
        std::filesystem::path output_path = input_path.parent_path() / input_path.stem();
        output_path += L"_converted.";
        output_path += to_lower(destination);
        return output_path;
    }

    bool is_file_valid_for_request(const std::filesystem::path& input_path)
    {
        std::error_code ec;
        if (!std::filesystem::exists(input_path, ec) || ec)
        {
            return false;
        }

        ec.clear();
        return std::filesystem::is_regular_file(input_path, ec) && !ec;
    }

    struct ConversionRequest
    {
        std::wstring destination;
        std::vector<std::wstring> files;
    };

    std::optional<ConversionRequest> parse_request_payload(const std::string& payload)
    {
        const nlohmann::json parsed = nlohmann::json::parse(payload, nullptr, false);
        if (parsed.is_discarded() || !parsed.is_object())
        {
            return std::nullopt;
        }

        if (!parsed.contains("action") || !parsed["action"].is_string())
        {
            return std::nullopt;
        }

        const std::string action = parsed["action"].get<std::string>();
        if (_stricmp(action.c_str(), ACTION_FORMAT_CONVERT) != 0)
        {
            return std::nullopt;
        }

        std::wstring destination = L"png";
        if (parsed.contains("destination") && parsed["destination"].is_string())
        {
            destination = utf8_to_wide(parsed["destination"].get<std::string>());
            if (destination.empty())
            {
                destination = L"png";
            }
        }

        if (!parsed.contains("files") || !parsed["files"].is_array())
        {
            return std::nullopt;
        }

        ConversionRequest request;
        request.destination = destination;

        for (const auto& file : parsed["files"])
        {
            if (!file.is_string())
=======
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
        constexpr DWORD buffer_size = 4096;
        char buffer[buffer_size] = {};
        std::string payload;

        while (true)
        {
            DWORD bytes_read = 0;
            const BOOL read_ok = ReadFile(pipe_handle, buffer, buffer_size, &bytes_read, nullptr);
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
>>>>>>> Stashed changes
            {
                continue;
            }

<<<<<<< Updated upstream
            std::wstring file_path = utf8_to_wide(file.get<std::string>());
            if (!file_path.empty())
            {
                request.files.push_back(std::move(file_path));
            }
        }

        if (request.files.empty())
        {
            return std::nullopt;
        }

        return request;
=======
            if (read_error == ERROR_BROKEN_PIPE)
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
        std::vector<std::wstring>& files,
        file_converter::ImageFormat& format,
        size_t& skipped_entries,
        std::wstring& rejection_reason)
    {
        files.clear();
        format = file_converter::ImageFormat::Png;
        skipped_entries = 0;
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
        if (_wcsicmp(action.c_str(), L"FormatConvert") != 0)
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
                ++skipped_entries;
                continue;
            }

            const auto file_path = file_value.GetString();
            if (file_path.empty())
            {
                ++skipped_entries;
                continue;
            }

            files.push_back(file_path.c_str());
        }

        if (files.empty())
        {
            rejection_reason = L"no valid file paths";
            return false;
        }

        format = parse_format(destination);
        return true;
    }

    ConversionSummary process_format_convert_request(const std::vector<std::wstring>& files, file_converter::ImageFormat format)
    {
        ConversionSummary summary;
        const std::wstring output_extension = extension_for_format(format);

        for (const auto& file : files)
        {
            const std::filesystem::path input_path(file);
            if (input_path.empty() || !std::filesystem::exists(input_path))
            {
                ++summary.missing_inputs;
                continue;
            }

            std::filesystem::path output_path = input_path.parent_path() / input_path.stem();
            output_path += L"_converted";
            output_path += output_extension;

            const auto conversion = file_converter::ConvertImageFile(input_path.wstring(), output_path.wstring(), format);
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
>>>>>>> Stashed changes
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

    class FileConverterPipeListener
    {
    public:
        void start()
        {
            if (m_running.exchange(true))
            {
                return;
            }

            m_pipe_name = get_pipe_name();
            m_worker_thread = std::thread(&FileConverterPipeListener::worker_loop, this);
            m_listener_thread = std::thread(&FileConverterPipeListener::listener_loop, this);
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
                std::swap(m_pending_requests, empty);
            }
        }

        ~FileConverterPipeListener()
        {
            stop();
        }

    private:
        void wake_listener()
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
                FILE_ATTRIBUTE_NORMAL,
                nullptr);

            if (wake_handle != INVALID_HANDLE_VALUE)
            {
                CloseHandle(wake_handle);
            }
        }

        void enqueue_request(std::string request)
        {
            {
                std::scoped_lock lock(m_queue_mutex);
                m_pending_requests.push(std::move(request));
            }

            m_queue_cv.notify_one();
        }

        HRESULT read_request(HANDLE pipe_handle, std::string& request)
        {
            char buffer[PIPE_READ_BUFFER_SIZE];

            while (m_running)
            {
                DWORD bytes_read = 0;
                const BOOL read_ok = ReadFile(pipe_handle, buffer, static_cast<DWORD>(sizeof(buffer)), &bytes_read, nullptr);
                if (read_ok)
                {
                    if (bytes_read > 0)
                    {
                        request.append(buffer, bytes_read);
                    }

                    break;
                }

                const DWORD read_error = GetLastError();
                if (read_error == ERROR_MORE_DATA)
                {
                    if (bytes_read > 0)
                    {
                        request.append(buffer, bytes_read);
                    }

                    continue;
                }

                if (read_error == ERROR_BROKEN_PIPE || read_error == ERROR_PIPE_NOT_CONNECTED)
                {
                    break;
                }

                return HRESULT_FROM_WIN32(read_error);
            }

            return request.empty() ? S_FALSE : S_OK;
        }

        void process_request(const std::string& request_payload)
        {
            const auto parsed_request = parse_request_payload(request_payload);
            if (!parsed_request.has_value())
            {
                Logger::error(L"File Converter: invalid pipe payload received");
                return;
            }

            const std::wstring destination = to_lower(parsed_request->destination);
            const auto format = parse_format(destination);
            std::unordered_set<std::wstring> seen_files;

            for (const auto& file_path : parsed_request->files)
            {
                const std::filesystem::path input_path(file_path);

                if (!seen_files.insert(input_path.wstring()).second)
                {
                    continue;
                }

                if (!is_file_valid_for_request(input_path))
                {
                    Logger::warn(L"File Converter skipped invalid input path '{}'", input_path.wstring());
                    continue;
                }

                const std::filesystem::path output_path = build_output_path(input_path, destination);

                const auto conversion = file_converter::ConvertImageFile(input_path.wstring(), output_path.wstring(), format);
                if (FAILED(conversion.hr))
                {
                    Logger::error(L"File Converter failed to convert '{}' to '{}'", input_path.wstring(), output_path.wstring());
                }
            }
        }

        void worker_loop()
        {
            while (m_running)
            {
                std::string request_payload;
                {
                    std::unique_lock lock(m_queue_mutex);
                    m_queue_cv.wait(lock, [this] {
                        return !m_running || !m_pending_requests.empty();
                    });

                    if (m_pending_requests.empty())
                    {
                        continue;
                    }

                    request_payload = std::move(m_pending_requests.front());
                    m_pending_requests.pop();
                }

                process_request(request_payload);
            }

            while (true)
            {
                std::string request_payload;
                {
                    std::scoped_lock lock(m_queue_mutex);
                    if (m_pending_requests.empty())
                    {
                        break;
                    }

                    request_payload = std::move(m_pending_requests.front());
                    m_pending_requests.pop();
                }

                process_request(request_payload);
            }
        }

        void listener_loop()
        {
            while (m_running)
            {
                HANDLE pipe_handle = CreateNamedPipeW(
                    m_pipe_name.c_str(),
                    PIPE_ACCESS_INBOUND | FILE_FLAG_OVERLAPPED,
                    PIPE_TYPE_MESSAGE | PIPE_READMODE_MESSAGE | PIPE_WAIT,
                    1,
                    0,
                    static_cast<DWORD>(PIPE_READ_BUFFER_SIZE),
                    0,
                    nullptr);

                if (pipe_handle == INVALID_HANDLE_VALUE)
                {
                    std::this_thread::sleep_for(std::chrono::milliseconds(250));
                    continue;
                }

                OVERLAPPED overlapped{};
                overlapped.hEvent = CreateEventW(nullptr, TRUE, FALSE, nullptr);
                if (overlapped.hEvent == nullptr)
                {
                    CloseHandle(pipe_handle);
                    std::this_thread::sleep_for(std::chrono::milliseconds(250));
                    continue;
                }

                const BOOL connect_result = ConnectNamedPipe(pipe_handle, &overlapped);
                const DWORD connect_error = connect_result ? ERROR_SUCCESS : GetLastError();
                if (!connect_result)
                {
                    if (connect_error == ERROR_PIPE_CONNECTED)
                    {
                        SetEvent(overlapped.hEvent);
                    }
                    else if (connect_error != ERROR_IO_PENDING)
                    {
                        CloseHandle(overlapped.hEvent);
                        CloseHandle(pipe_handle);
                        continue;
                    }
                }

                bool connected = false;
                while (m_running)
                {
                    const DWORD wait_result = WaitForSingleObject(overlapped.hEvent, 250);
                    if (wait_result == WAIT_OBJECT_0)
                    {
                        connected = true;
                        break;
                    }

                    if (wait_result == WAIT_FAILED)
                    {
                        break;
                    }
                }

                if (!m_running)
                {
                    CancelIoEx(pipe_handle, &overlapped);
                    CloseHandle(overlapped.hEvent);
                    CloseHandle(pipe_handle);
                    break;
                }

                if (connected)
                {
                    std::string request_payload;
                    if (SUCCEEDED(read_request(pipe_handle, request_payload)) && !request_payload.empty())
                    {
                        enqueue_request(std::move(request_payload));
                    }
                }

                DisconnectNamedPipe(pipe_handle);
                CloseHandle(overlapped.hEvent);
                CloseHandle(pipe_handle);
            }
        }

        std::atomic<bool> m_running = false;
        std::wstring m_pipe_name;
        std::thread m_listener_thread;
        std::thread m_worker_thread;
        std::mutex m_queue_mutex;
        std::condition_variable m_queue_cv;
        std::queue<std::string> m_pending_requests;
    };
}

class FileConverterModule : public PowertoyModuleIface
{
public:
<<<<<<< Updated upstream
    FileConverterModule()
    {
        m_pipe_listener.start();
=======
    ~FileConverterModule()
    {
        disable();
>>>>>>> Stashed changes
    }

    void destroy() override
    {
        m_pipe_listener.stop();
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

    void call_custom_action(const wchar_t* /*action*/) override
    {
<<<<<<< Updated upstream
=======
        if (action == nullptr)
        {
            return;
        }

        handle_payload(winrt::to_string(action));
>>>>>>> Stashed changes
    }

    void enable() override
    {
        if (m_enabled)
        {
            return;
        }

        ensure_context_menu_package_registered();
        ensure_context_menu_runtime_registered();
<<<<<<< Updated upstream
        m_pipe_listener.start();
=======

        m_pipe_name = get_pipe_name_for_current_session();
        m_stop_requested.store(false);
        m_listener_thread = std::thread([this]() {
            run_listener();
        });

>>>>>>> Stashed changes
        m_enabled = true;
    }

    void disable() override
    {
<<<<<<< Updated upstream
        m_pipe_listener.stop();
=======
        if (!m_enabled && !m_listener_thread.joinable())
        {
            return;
        }

        m_stop_requested.store(true);
        signal_listener_shutdown();

        if (m_listener_thread.joinable())
        {
            m_listener_thread.join();
        }

>>>>>>> Stashed changes
        unregister_context_menu_runtime();
        m_enabled = false;
    }

    bool is_enabled() override
    {
        return m_enabled;
    }

private:
    void handle_payload(const std::string& payload)
    {
        std::vector<std::wstring> files;
        file_converter::ImageFormat format = file_converter::ImageFormat::Png;
        size_t skipped_entries = 0;
        std::wstring rejection_reason;
        if (!try_parse_format_convert_request(payload, files, format, skipped_entries, rejection_reason))
        {
            if (!rejection_reason.empty())
            {
                Logger::warn(L"File Converter ignored malformed request: {}", rejection_reason);
            }

            return;
        }

        const auto summary = process_format_convert_request(files, format);

        if (skipped_entries > 0)
        {
            Logger::warn(L"File Converter request skipped {} invalid file entries.", skipped_entries);
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

    void run_listener()
    {
        while (!m_stop_requested.load())
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
                if (!m_stop_requested.load())
                {
                    std::this_thread::sleep_for(std::chrono::milliseconds(50));
                }

                continue;
            }

            const std::string payload = read_pipe_message(pipe_handle);

            FlushFileBuffers(pipe_handle);
            DisconnectNamedPipe(pipe_handle);
            CloseHandle(pipe_handle);

            if (m_stop_requested.load())
            {
                break;
            }

            handle_payload(payload);
        }
    }

    void signal_listener_shutdown() const
    {
        if (m_pipe_name.empty())
        {
            return;
        }

        if (!WaitNamedPipeW(m_pipe_name.c_str(), 100))
        {
            return;
        }

        HANDLE pipe_handle = CreateFileW(
            m_pipe_name.c_str(),
            GENERIC_WRITE,
            0,
            nullptr,
            OPEN_EXISTING,
            0,
            nullptr);

        if (pipe_handle == INVALID_HANDLE_VALUE)
        {
            return;
        }

        constexpr char shutdown_message[] = "{}";
        DWORD bytes_written = 0;
        (void)WriteFile(pipe_handle, shutdown_message, sizeof(shutdown_message) - 1, &bytes_written, nullptr);
        CloseHandle(pipe_handle);
    }

    bool m_enabled = false;
<<<<<<< Updated upstream
    FileConverterPipeListener m_pipe_listener;
=======
    std::atomic<bool> m_stop_requested = false;
    std::thread m_listener_thread;
    std::wstring m_pipe_name;
>>>>>>> Stashed changes
};

extern "C" __declspec(dllexport) PowertoyModuleIface* __cdecl powertoy_create()
{
    return new FileConverterModule();
}
