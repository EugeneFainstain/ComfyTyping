// ComfyTyping.cpp : Defines the entry point for the application.
//

#include "framework.h"
#include "ComfyTyping.h"

#define MAX_LOADSTRING 100

// Global Variables:
HINSTANCE hInst;                                // current instance
WCHAR szTitle[MAX_LOADSTRING];                  // The title bar text
WCHAR szWindowClass[MAX_LOADSTRING];            // the main window class name

#define INVALIDATE_ON_TIMER    2
#define INVALIDATE_ON_HOOK     3
#define INVALIDATE_METHOD   INVALIDATE_ON_HOOK //INVALIDATE_ON_TIMER //3  

HWND g_myWindowHandle = nullptr;
HWINEVENTHOOK g_hEventHook = nullptr;

// Forward declarations of functions included in this code module:
ATOM                MyRegisterClass(HINSTANCE hInstance);
HWND                InitInstance(HINSTANCE, int);
LRESULT CALLBACK    WndProc(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK    About(HWND, UINT, WPARAM, LPARAM);

#if INVALIDATE_METHOD == INVALIDATE_ON_TIMER
    #define TIMER_ID       1001  // Unique timer ID
    #define TIMER_INTERVAL 16    // 16 ms interval
#endif

int APIENTRY wWinMain(_In_ HINSTANCE hInstance,
                     _In_opt_ HINSTANCE hPrevInstance,
                     _In_ LPWSTR    lpCmdLine,
                     _In_ int       nCmdShow)
{
    UNREFERENCED_PARAMETER(hPrevInstance);
    UNREFERENCED_PARAMETER(lpCmdLine);

    SetProcessDPIAware();  // Enable DPI awareness (Windows 7+)

    // TODO: Place code here.

    // Initialize global strings
    LoadStringW(hInstance, IDS_APP_TITLE, szTitle, MAX_LOADSTRING);
    LoadStringW(hInstance, IDC_COMFYTYPING, szWindowClass, MAX_LOADSTRING);
    MyRegisterClass(hInstance);

    // Perform application initialization:
    g_myWindowHandle = InitInstance(hInstance, nCmdShow);

    if (!g_myWindowHandle)
    {
        return FALSE;
    }

    HACCEL hAccelTable = LoadAccelerators(hInstance, MAKEINTRESOURCE(IDC_COMFYTYPING));

#if INVALIDATE_METHOD == INVALIDATE_ON_TIMER
    SetTimer(g_myWindowHandle, TIMER_ID, TIMER_INTERVAL, NULL); // Start the timer
#endif

    MSG msg;

    // Main message loop:
    while (GetMessage(&msg, nullptr, 0, 0))
    {
        if (!TranslateAccelerator(msg.hwnd, hAccelTable, &msg))
        {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }

    return (int) msg.wParam;
}

#if INVALIDATE_METHOD == INVALIDATE_ON_HOOK
void CALLBACK WinEventProc(HWINEVENTHOOK hWinEventHook, DWORD event, HWND hwnd,
                           LONG idObject, LONG idChild, DWORD dwEventThread, DWORD dwmsEventTime) {
    if (hwnd != g_myWindowHandle) {  // Ignore own window
        InvalidateRect(g_myWindowHandle, NULL, FALSE);
    }
}

void StartDesktopMonitor() {
    g_hEventHook = SetWinEventHook(EVENT_OBJECT_LOCATIONCHANGE, EVENT_OBJECT_NAMECHANGE,
                                 NULL, WinEventProc, 0, 0, WINEVENT_OUTOFCONTEXT);
}

void StopDesktopMonitor() {
    if (g_hEventHook) UnhookWinEvent(g_hEventHook);
}
#endif // INVALIDATE_METHOD == INVALIDATE_ON_HOOK

//
//  FUNCTION: MyRegisterClass()
//
//  PURPOSE: Registers the window class.
//
ATOM MyRegisterClass(HINSTANCE hInstance)
{
    WNDCLASSEXW wcex;

    wcex.cbSize = sizeof(WNDCLASSEX);

    wcex.style          = CS_HREDRAW | CS_VREDRAW;
    wcex.lpfnWndProc    = WndProc;
    wcex.cbClsExtra     = 0;
    wcex.cbWndExtra     = 0;
    wcex.hInstance      = hInstance;
    wcex.hIcon          = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_COMFYTYPING));
    wcex.hCursor        = LoadCursor(nullptr, IDC_ARROW);
    wcex.hbrBackground  = (HBRUSH)(COLOR_WINDOW+1);
    wcex.lpszMenuName   = MAKEINTRESOURCEW(IDC_COMFYTYPING);
    wcex.lpszClassName  = szWindowClass;
    wcex.hIconSm        = LoadIcon(wcex.hInstance, MAKEINTRESOURCE(IDI_SMALL));

    return RegisterClassExW(&wcex);
}

//
//   FUNCTION: InitInstance(HINSTANCE, int)
//
//   PURPOSE: Saves instance handle and creates main window
//
//   COMMENTS:
//
//        In this function, we save the instance handle in a global variable and
//        create and display the main program window.
//
HWND InitInstance(HINSTANCE hInstance, int nCmdShow)
{
   hInst = hInstance; // Store instance handle in our global variable

   HWND hWnd = CreateWindowExW(WS_EX_COMPOSITED, szWindowClass, szTitle, WS_OVERLAPPEDWINDOW,
      CW_USEDEFAULT, 0, CW_USEDEFAULT, 0, nullptr, nullptr, hInstance, nullptr);

   if (!hWnd)
   {
      return hWnd;
   }

   ShowWindow(hWnd, nCmdShow);
   UpdateWindow(hWnd);

   // Make topmost
   SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);

#if INVALIDATE_METHOD == INVALIDATE_ON_HOOK
   // Start monitoring desktop for changes
   StartDesktopMonitor();
#endif

   return hWnd;
}

void RenderScaledCursor(HDC hTargetDC, HWND hWnd, int width, int height)
{
    // Get cursor info
    CURSORINFO ci;
    ci.cbSize = sizeof(CURSORINFO);
    if (GetCursorInfo(&ci) && ci.flags == CURSOR_SHOWING)
    {
        // Get DPI scaling factor
        int dpi = GetDpiForWindow(hWnd);  // Windows 8.1+ (Use 96 DPI as base)
        float scaleFactor = dpi / 96.0f;

        // Get the cursor bitmap
        ICONINFO iconInfo;
        GetIconInfo(ci.hCursor, &iconInfo);

        // Get original cursor size
        BITMAP bmp;
        GetObject(iconInfo.hbmMask, sizeof(BITMAP), &bmp);
        int cursorWidth = bmp.bmWidth;
        int cursorHeight = bmp.bmHeight;

        // Scale the cursor size
        int scaledWidth = (int)(cursorWidth * scaleFactor);
        int scaledHeight = (int)(cursorHeight * scaleFactor);

        // Adjust cursor position based on the hotspot
        int hotspotX = (int)(iconInfo.xHotspot * scaleFactor);
        int hotspotY = (int)(iconInfo.yHotspot * scaleFactor);

        // Convert cursor screen position to local coordinates, accounting for the hotspot
        POINT cursorPos = ci.ptScreenPos;
        int drawX = cursorPos.x - hotspotX;
        int drawY = cursorPos.y - hotspotY;

        // Draw the scaled cursor
        DrawIconEx(hTargetDC, drawX, drawY, ci.hCursor,
                   scaledWidth, scaledHeight, 0, NULL, DI_NORMAL);

        // Clean up
        DeleteObject(iconInfo.hbmMask);
        DeleteObject(iconInfo.hbmColor);
    }
}

//
//  FUNCTION: WndProc(HWND, UINT, WPARAM, LPARAM)
//
//  PURPOSE: Processes messages for the main window.
//
//  WM_COMMAND  - process the application menu
//  WM_PAINT    - Paint the main window
//  WM_DESTROY  - post a quit message and return
//
//
LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    bool bDestroy = false;

    switch (message)
    {
    case WM_COMMAND:
        {
            int wmId = LOWORD(wParam);
            // Parse the menu selections:
            switch (wmId)
            {
            case IDM_ABOUT:
                DialogBox(hInst, MAKEINTRESOURCE(IDD_ABOUTBOX), hWnd, About);
                break;
            case IDM_EXIT:
                DestroyWindow(hWnd);
                break;
            default:
                return DefWindowProc(hWnd, message, wParam, lParam);
            }
        }
        break;
#if INVALIDATE_METHOD == INVALIDATE_ON_TIMER
    case WM_TIMER:
        if (wParam == TIMER_ID) // Ensure it's our specific timer
        {
            InvalidateRect(hWnd, NULL, FALSE);
        }
        break;
#endif
    case WM_MOVING:
    case WM_SIZING:
        ValidateRect(hWnd, NULL);
        break;
    case WM_PAINT:
        {
            PAINTSTRUCT ps;
            HDC hdc = BeginPaint(hWnd, &ps);
            // TODO: Add any drawing code that uses hdc here...
            {
                HDC hScreenDC = GetDC(NULL); // Get the desktop DC
                HDC hMemDC = CreateCompatibleDC(hScreenDC);

                int width = 600, height = 400; // Define the region to capture
                HBITMAP hBitmap = CreateCompatibleBitmap(hScreenDC, width, height);
                SelectObject(hMemDC, hBitmap);

                // Copy from screen to memory DC
                BitBlt(hMemDC, 0, 0, width, height, hScreenDC, 0, 0, SRCCOPY); //GetSystemMetrics(SM_CXSCREEN), GetSystemMetrics(SM_CYSCREEN),

                ReleaseDC(NULL, hScreenDC);

                RenderScaledCursor(hMemDC, hWnd, width, height); // Render to hMemDC

                BitBlt(hdc, 0, 0, width, height, hMemDC, 0, 0, SRCCOPY);

                // RenderScaledCursor(hdc, hWnd, width, height); // Render directly to hdc - this also works

                DeleteDC(hMemDC);
            }
            EndPaint(hWnd, &ps);
        }
        break;
    case WM_DESTROY:
        bDestroy = true;
        break;
    default:
        return DefWindowProc(hWnd, message, wParam, lParam);
    }
    //////////////////////////////////////////////

    if (bDestroy)
    {
        #if INVALIDATE_METHOD == INVALIDATE_ON_TIMER
            KillTimer(hWnd, TIMER_ID); // Kill the timer
        #endif

        #if INVALIDATE_METHOD == INVALIDATE_ON_HOOK
            StopDesktopMonitor(); // Stop monitoring desktop for changes
        #endif

        PostQuitMessage(0);
    }

    return 0;
}

// Message handler for about box.
INT_PTR CALLBACK About(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
    UNREFERENCED_PARAMETER(lParam);
    switch (message)
    {
    case WM_INITDIALOG:
        return (INT_PTR)TRUE;

    case WM_COMMAND:
        if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL)
        {
            EndDialog(hDlg, LOWORD(wParam));
            return (INT_PTR)TRUE;
        }
        break;
    }
    return (INT_PTR)FALSE;
}
