#include "pch.h"

#include <interface/powertoy_module_interface.h>

namespace
{
    constexpr wchar_t MODULE_NAME[] = L"File Converter";
    constexpr wchar_t MODULE_KEY[] = L"FileConverter";
    constexpr wchar_t MODULE_CONFIG[] = LR"({"name":"File Converter","version":"1.0","properties":{}})";
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

    void enable() override
    {
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
