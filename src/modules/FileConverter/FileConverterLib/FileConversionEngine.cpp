#include "pch.h"

#include "FileConversionEngine.h"

#include <wrl/client.h>

namespace
{
    GUID container_format_for(file_converter::ImageFormat format)
    {
        switch (format)
        {
        case file_converter::ImageFormat::Jpeg:
            return GUID_ContainerFormatJpeg;
        case file_converter::ImageFormat::Bmp:
            return GUID_ContainerFormatBmp;
        case file_converter::ImageFormat::Tiff:
            return GUID_ContainerFormatTiff;
        case file_converter::ImageFormat::Png:
        default:
            return GUID_ContainerFormatPng;
        }
    }

    std::wstring hr_message(const wchar_t* prefix, HRESULT hr)
    {
        return std::wstring(prefix) + L" HRESULT=0x" + std::to_wstring(static_cast<unsigned long>(hr));
    }

    struct ScopedCom
    {
        HRESULT hr;

        ScopedCom()
            : hr(CoInitializeEx(nullptr, COINIT_MULTITHREADED))
        {
        }

        ~ScopedCom()
        {
            if (SUCCEEDED(hr))
            {
                CoUninitialize();
            }
        }
    };
}

namespace file_converter
{
    ConversionResult ConvertImageFile(const std::wstring& input_path, const std::wstring& output_path, ImageFormat format)
    {
        ScopedCom com;
        if (FAILED(com.hr))
        {
            return { com.hr, hr_message(L"CoInitializeEx failed.", com.hr) };
        }

        Microsoft::WRL::ComPtr<IWICImagingFactory> factory;
        HRESULT hr = CoCreateInstance(CLSID_WICImagingFactory2, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&factory));
        if (FAILED(hr))
        {
            hr = CoCreateInstance(CLSID_WICImagingFactory, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&factory));
        }

        if (FAILED(hr))
        {
            return { hr, hr_message(L"Failed creating WIC factory.", hr) };
        }

        Microsoft::WRL::ComPtr<IWICBitmapDecoder> decoder;
        hr = factory->CreateDecoderFromFilename(input_path.c_str(), nullptr, GENERIC_READ, WICDecodeMetadataCacheOnLoad, &decoder);
        if (FAILED(hr))
        {
            return { hr, hr_message(L"Failed opening input image.", hr) };
        }

        Microsoft::WRL::ComPtr<IWICBitmapFrameDecode> source_frame;
        hr = decoder->GetFrame(0, &source_frame);
        if (FAILED(hr))
        {
            return { hr, hr_message(L"Failed reading first image frame.", hr) };
        }

        UINT width = 0;
        UINT height = 0;
        hr = source_frame->GetSize(&width, &height);
        if (FAILED(hr))
        {
            return { hr, hr_message(L"Failed reading image size.", hr) };
        }

        WICPixelFormatGUID pixel_format = {};
        hr = source_frame->GetPixelFormat(&pixel_format);
        if (FAILED(hr))
        {
            return { hr, hr_message(L"Failed reading source pixel format.", hr) };
        }

        Microsoft::WRL::ComPtr<IWICStream> output_stream;
        hr = factory->CreateStream(&output_stream);
        if (FAILED(hr))
        {
            return { hr, hr_message(L"Failed creating WIC stream.", hr) };
        }

        hr = output_stream->InitializeFromFilename(output_path.c_str(), GENERIC_WRITE);
        if (FAILED(hr))
        {
            return { hr, hr_message(L"Failed opening output path.", hr) };
        }

        Microsoft::WRL::ComPtr<IWICBitmapEncoder> encoder;
        hr = factory->CreateEncoder(container_format_for(format), nullptr, &encoder);
        if (FAILED(hr))
        {
            return { hr, hr_message(L"Failed creating image encoder.", hr) };
        }

        hr = encoder->Initialize(output_stream.Get(), WICBitmapEncoderNoCache);
        if (FAILED(hr))
        {
            return { hr, hr_message(L"Failed initializing encoder.", hr) };
        }

        Microsoft::WRL::ComPtr<IWICBitmapFrameEncode> target_frame;
        Microsoft::WRL::ComPtr<IPropertyBag2> frame_properties;
        hr = encoder->CreateNewFrame(&target_frame, &frame_properties);
        if (FAILED(hr))
        {
            return { hr, hr_message(L"Failed creating target frame.", hr) };
        }

        hr = target_frame->Initialize(frame_properties.Get());
        if (FAILED(hr))
        {
            return { hr, hr_message(L"Failed initializing target frame.", hr) };
        }

        hr = target_frame->SetSize(width, height);
        if (FAILED(hr))
        {
            return { hr, hr_message(L"Failed setting target size.", hr) };
        }

        hr = target_frame->SetPixelFormat(&pixel_format);
        if (FAILED(hr))
        {
            return { hr, hr_message(L"Failed setting target pixel format.", hr) };
        }

        hr = target_frame->WriteSource(source_frame.Get(), nullptr);
        if (FAILED(hr))
        {
            return { hr, hr_message(L"Failed writing target frame.", hr) };
        }

        hr = target_frame->Commit();
        if (FAILED(hr))
        {
            return { hr, hr_message(L"Failed committing target frame.", hr) };
        }

        hr = encoder->Commit();
        if (FAILED(hr))
        {
            return { hr, hr_message(L"Failed committing encoder.", hr) };
        }

        return { S_OK, L"" };
    }
}
