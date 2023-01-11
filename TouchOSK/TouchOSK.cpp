#include <iostream>
#include <initguid.h>
#include <Objbase.h>
#include <Shobjidl.h>

#pragma hdrstop


// 4ce576fa-83dc-4F88-951c-9d0782b4e376
DEFINE_GUID(CLSID_UIHostNoLaunch,
    0x4CE576FA, 0x83DC, 0x4f88, 0x95, 0x1C, 0x9D, 0x07, 0x82, 0xB4, 0xE3, 0x76);

// 37c994e7_432b_4834_a2f7_dce1f13b834b
DEFINE_GUID(IID_ITipInvocation,
    0x37c994e7, 0x432b, 0x4834, 0xa2, 0xf7, 0xdc, 0xe1, 0xf1, 0x3b, 0x83, 0x4b);

struct ITipInvocation : IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE Toggle(HWND wnd) = 0;
};

using namespace std;

/// <summary>
/// Determines whether an input pane (Touch Keyboard or handwriting Panel) is currently open.
/// </summary>
/// <returns>
/// If the function succeeded and an input pane is currently open, S_OK is retured.
/// If the function succeeded and no input pane is currently open, S_FALSE is returnd.
/// 
/// Otherwise, another error code is returned.
/// </returns>
HRESULT IsInputPaneOpen()
{
    RECT rect;
    ZeroMemory(&rect, sizeof(rect));

    IFrameworkInputPane* frameworkInputPane{ nullptr };
    HRESULT hr{ CoCreateInstance(CLSID_FrameworkInputPane, NULL, CLSCTX_INPROC_SERVER, IID_IFrameworkInputPane, (void**)&frameworkInputPane) };
    if (SUCCEEDED(hr))
    {
        hr = frameworkInputPane->Location(&rect);
        if (SUCCEEDED(hr))
        {
            hr = IsRectEmpty(&rect) ? S_FALSE : S_OK;
        }
        frameworkInputPane->Release();
    }

    return hr;
}

int WinMain(HINSTANCE hInst, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow)
{
    HRESULT hr{ CoInitialize(NULL) };
    if (FAILED(hr))
    {
        wcerr << L"Failed to initialize COM." << endl;
        return 1;
    }

    // Toggle the Touch Keyboard regardless of whether it is currently shown or not.
    // To only show the Touch Keyboard and not hide it if it is already active,
    // abort here if IsInputPaneOpen() == S_OK.

    ITipInvocation* tip{ nullptr };
    hr = CoCreateInstance(CLSID_UIHostNoLaunch, NULL, CLSCTX_INPROC_HANDLER | CLSCTX_LOCAL_SERVER, IID_ITipInvocation, (void**)&tip);

    if (hr == REGDB_E_CLASSNOTREG)
    {
        INT_PTR result = (INT_PTR)ShellExecuteW(NULL, NULL, L"TabTip.exe", NULL, NULL, SW_SHOWNORMAL);
        if (result > 32)
        {
            wcout << L"Started TabTip.exe to open Touch Keyboard." << endl;
        }
        else
        {
            wcerr << L"Failed to start TabTip.exe. Error: " << result << endl;
        }
    }
    else if (SUCCEEDED(hr))
    {
        HWND desktopWindow = GetDesktopWindow();
        hr = tip->Toggle(desktopWindow);
        if (SUCCEEDED(hr))
        {
            wcout << L"Toggled the touch keyboard via ITipInvocation.Toggle()." << endl;
        }
        else
        {
            wcerr << L"Failed to toggle the Touch Keyboard via ITipInvocation.Toggle()." << endl;
        }
        tip->Release();
    }

    CoUninitialize();
}