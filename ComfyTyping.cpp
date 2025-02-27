// ComfyTyping.cpp : Defines the entry point for the application.
//

#define WIN_UTILS_CPP

#include "framework.h"
#include "resource.h"
#include "ComfyTyping.h"
#include "WinUtils.h"
#include <iostream> // for std::isprint

#define MAX_LOADSTRING 100

// Global Variables:
HINSTANCE hInst;                                // current instance
WCHAR szTitle[MAX_LOADSTRING];                  // The title bar text
WCHAR szWindowClass[MAX_LOADSTRING];            // the main window class name

void CALLBACK WinEventProc_ForFocusedClientWnd(HWINEVENTHOOK hWinEventHook, DWORD event, HWND hwnd,
                           LONG idObject, LONG idChild, DWORD dwEventThread, DWORD dwmsEventTime);

LRESULT CALLBACK LowLevelKeyboardProc(int nCode, WPARAM wParam, LPARAM lParam);

// Forward declarations of functions included in this code module:
ATOM                MyRegisterClass(HINSTANCE hInstance);
HWND                InitInstance(HINSTANCE, int);
LRESULT CALLBACK    WndProc(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK    About(HWND, UINT, WPARAM, LPARAM);

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

#ifdef INVALIDATE_ON_TIMER
    SetTimer(g_myWindowHandle, INVALIDATE_TIMER_ID, INVALIDATE_TIMER_INTERVAL, NULL); // Start the timer
#endif

    SetTimer(g_myWindowHandle, CARET_TIMER_ID, CARET_TIMER_INTERVAL, NULL); // Start the timer

    HWINEVENTHOOK hWinEventHook = SetWinEventHook(
                                    EVENT_OBJECT_FOCUS,// [in] DWORD      eventMin,
                                    EVENT_OBJECT_FOCUS,// [in] DWORD      eventMax,
                                    NULL,              // [in] HMODULE    hmodWinEventProc,
                                    WinEventProc_ForFocusedClientWnd,      // [in] WINEVENTPROC pfnWinEventProc,
                                    0,                 // [in] DWORD        idProcess,
                                    0,                 // [in] DWORD        idThread,
                                    WINEVENT_OUTOFCONTEXT | WINEVENT_SKIPOWNPROCESS | WINEVENT_SKIPOWNTHREAD );//[in] DWORD        dwFlags

    HHOOK hHook = SetWindowsHookEx(WH_KEYBOARD_LL, LowLevelKeyboardProc, hInstance, 0);

    if( !hHook || !hWinEventHook )
    {
        ::MessageBoxA(0, "Error: please re-run as an Administrator.\n", "ComfyTyping", MB_ICONERROR);
        exit(-1);
    }

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

#ifdef INVALIDATE_ON_HOOK
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
#endif // INVALIDATE_ON_HOOK

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

#ifdef NO_RESIZE_NO_BORDER
    HWND hWnd = CreateWindowExW(WS_EX_COMPOSITED, szWindowClass, szTitle, WS_POPUP | WS_MINIMIZEBOX | WS_SYSMENU, //WS_OVERLAPPEDWINDOW,
       CW_USEDEFAULT, 0, CW_USEDEFAULT, 0, nullptr, nullptr, hInstance, nullptr);

    EnableRoundedCorners(hWnd); // Enable rounded corners for the popup window
#else
    HWND hWnd = CreateWindowExW(WS_EX_COMPOSITED, szWindowClass, szTitle, WS_OVERLAPPEDWINDOW, //WS_POPUP | WS_MINIMIZEBOX | WS_SYSMENU,
       CW_USEDEFAULT, 0, CW_USEDEFAULT, 0, nullptr, nullptr, hInstance, nullptr);
#endif

    if (!hWnd)
    {
       return hWnd;
    }

#if !defined(HAS_MENU)
    SetMenu(hWnd, nullptr);
#endif

#if !defined(HAS_TITLE_BAR)
    LONG_PTR style = GetWindowLongPtr(hWnd, GWL_STYLE); // Retrieve the current window style
    style &= ~WS_CAPTION; // Remove the title bar (caption) from the style
    SetWindowLongPtr(hWnd, GWL_STYLE, style); // Set the new style
    //SetWindowPos(hWnd, NULL, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED); // Apply the style changes
#endif

    ShowWindow(hWnd, nCmdShow);
    UpdateWindow(hWnd);

    // Place and make topmost
//    SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);
    int titleBarHeight = GetSystemMetrics(SM_CYCAPTION);
    int menuBarHeight  = GetSystemMetrics(SM_CYMENU);
    int borderHeight   = GetSystemMetrics(SM_CYSIZEFRAME);
    g_fDpiScaleFactor  = GetDpiScaleFactor(hWnd);
    g_iScreenWidth     = GetSystemMetrics(SM_CXSCREEN);
    g_iScreenHeight    = GetSystemMetrics(SM_CYSCREEN);

    g_iMyWidth         = g_iScreenWidth - g_iScreenWidth / HORZ_PORTION;
    g_iMyHeight        = g_iScreenHeight * ZOOM / VERT_PORTION;
#if !defined(NO_RESIZE_NO_BORDER)
    g_iMyHeight += borderHeight*4;
#endif
#ifdef HAS_MENU
    height += menuBarHeight * g_fDpiScaleFactor;
#endif
#ifdef HAS_TITLE_BAR
    height += (int)(titleBarHeight * g_fDpiScaleFactor);
#endif
    g_iMyWindowX  = (g_iScreenWidth - g_iMyWidth) / 2;
    g_iMyWindowY  = g_iScreenHeight - g_iMyHeight;
#if !defined(POSITION_OVER_TASKBAR)
    g_iMyWindowY -= GetTaskbarHeight();
#endif
#if !defined(NO_RESIZE_NO_BORDER)
    g_iMyWindowY += borderHeight*2;
#endif
    SetWindowPos(hWnd, HWND_TOPMOST, g_iMyWindowX, g_iMyWindowY, g_iMyWidth, g_iMyHeight, SWP_SHOWWINDOW);

#ifdef INVALIDATE_ON_HOOK
    // Start monitoring desktop for changes
    StartDesktopMonitor();
#endif

    return hWnd;
}

//bool g_bWindowAtTheBottom = false;
void MySetWindowPos(HWND hWnd, bool bShow)
     { SetWindowPos(hWnd, HWND_TOPMOST, g_iMyWindowX, (bShow ? 1 : 4) * g_iMyWindowY, 0, 0, SWP_NOSIZE | SWP_SHOWWINDOW | SWP_NOZORDER | SWP_NOACTIVATE); }

// void HideMyWindow(HWND hWnd) { if(!g_bWindowAtTheBottom) MySetWindowPos(hWnd, false); g_bWindowAtTheBottom = true; }
// void ShowMyWindow(HWND hWnd) {                           MySetWindowPos(hWnd, true ); g_bWindowAtTheBottom = false; }
void HideMyWindow(HWND hWnd) { MySetWindowPos(hWnd, false); }
void ShowMyWindow(HWND hWnd) { MySetWindowPos(hWnd, true ); }
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

    // Hiding the window on Esc, re-showing on mouse click:
    // Note: we can probably move this logic to LowLevelKeyboardProc()...
    DO_WHEN_BECOMES_TRUE( KEY_DOWN(VK_ESCAPE), g_bTemporarilyHideMyWindow = !g_bTemporarilyHideMyWindow);

    // Handling the messages
    switch (message)
    {
    case WM_SYSCOMMAND:
        if (wParam == SC_MINIMIZE) // Blocking the "minimize" functionality...
        {
//             if( g_bWindowAtTheBottom )
//                 ShowMyWindow(hWnd); // Toggle Show/Hide instead of restore...
//             else
//                 HideMyWindow(hWnd); // Toggle Show/Hide instead of minimization...

            //g_bTemporarilyHideMyWindow = !g_bTemporarilyHideMyWindow;
            return (LPARAM)0; // Not calling DefWindowProc() here
        }
        break;
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
            }
        }
        break;
    case WM_TIMER:
#ifdef INVALIDATE_ON_TIMER
        if (wParam == INVALIDATE_TIMER_ID) // Ensure it's our specific timer
        {
            InvalidateRect(hWnd, NULL, FALSE);
        }
#endif
        if( wParam == CARET_TIMER_ID ) // Ensure it's our specific timer
        if( g_hForegroundWindow = GetForegroundWindow() ) // g_hForegroundWindow can be NULL during focus transitions
        {
            g_fDpiScaleFactor = GetDpiScaleFactor(g_hForegroundWindow); // Update DPI scaling factor

            POINT caret = GetCaretPosition();

            bool bCaretFromAccessibility = false;
            if (caret.y == 0)
            {
                caret = GetCaretPositionFromAccessibility(); // this is too slow to be executed on WM_PAINT...
                bCaretFromAccessibility = true;
            }
            g_bCaretFromAccessibility = bCaretFromAccessibility;

            if ((caret.y < 0) || (caret.y > g_iScreenHeight))
                caret.x = caret.y = 0;

            static bool bWindowAtTheBottom = false;

            if( ((caret.y != 0) && (g_bTemporarilyHideMyWindow == false)) || g_hForegroundWindow == hWnd )
                ShowMyWindow(hWnd); // Make sure we are still topmost (needed to be above the taskbar)
            else
                HideMyWindow(hWnd); // This is now getting called on every WM_TIMER event... hmmm...

            if( g_hForegroundWindow != hWnd ) // Don't update the caret position if our window is in focus...
                g_ptCaret = caret;
        }
        break;
    case WM_MOVING:
    case WM_SIZING:
        ValidateRect(hWnd, NULL);
        break;
    case WM_PAINT:
        {
            PAINTSTRUCT ps;
            HDC hdc = BeginPaint(hWnd, &ps);

            if( !g_hFocusedChildWnd )
                g_hFocusedChildWnd = g_hForegroundWindow;

            if( g_hFocusedChildWnd )
            {
                int width = g_iMyWidth, height = g_iScreenHeight / VERT_PORTION; // Define the region to             capture

                POINT ptCaret = g_ptCaret; // Calling GetCaretPositionFromAccessibility() here is too slow...

                RECT focusedChildRect;
                HWND hFocusedChildWnd = g_hFocusedChildWnd;
                if( g_bCaretFromAccessibility && GetParent(hFocusedChildWnd) )
                    hFocusedChildWnd = GetParent(hFocusedChildWnd);
                GetWindowRect(hFocusedChildWnd, &focusedChildRect); // Should be always valid...

                int iSrcX = 0;
                int iSrcY = 0;

                if( ptCaret.y != 0 )
                {
                    #ifdef USE_FONT_HEIGHT
                        int iFontHeight = GetFontHeight(g_hForegroundWindow);
                    #endif
                    //iCaretY -= iFontHeight * 2;

                    //iSrcX = ptCaret.x / ZOOM;             // Cursor always matches real cursor
                    int iSrcX_left   = focusedChildRect.left;                      // Cursor starts from the left edge
                    int iSrcX_right  = focusedChildRect.right - g_iMyWidth / ZOOM; // Cursor starts from the right edge
                    //iSrcX = width / ZOOM / 2;             // Cursor center matches screen center
                    int iSrcX_center = ptCaret.x - g_iMyWidth / ZOOM / 2; // Cursor always in the center of the screen

                    // Implement the left-center-right logic
                    if( iSrcX_left >= iSrcX_right ) // This happens when the window width is small enough to fit as a whole
                        iSrcX = (iSrcX_left + iSrcX_right) / 2; //iSrcX = iSrcX_center;
                    else
                        iSrcX = max( iSrcX_left, min( iSrcX_right, iSrcX_center) );
                    //iSrcX = iSrcX_left;
                    //iSrcX = iSrcX_right;
                    //iSrcX = iSrcX_center;
                    iSrcY = ptCaret.y - height/ 2;
                }

                if( g_bFreezeSrcXandSrcY || g_iTakeCaretSnapshot )
                {
                    if( g_iTakeCaretSnapshot )
                    {
                        g_ptCaretSnapshot = g_ptCaret;
                        g_iTakeCaretSnapshot--;
                    }
                    else
                    if( g_ptCaret.x != g_ptCaretSnapshot.x || g_ptCaret.y != g_ptCaretSnapshot.y )
                        g_bFreezeSrcXandSrcY = false;
                }
                else
                    g_iSrcX = iSrcX;

                g_iSrcY = iSrcY; // Y is always updated...

                #ifdef RENDER_CURSOR
                    RenderScaledCursor(hMemDC, hWnd, width, height); // Render to hMemDC
                #endif

                // Copy a portion of the desktop image to the window
                {
                    HDC hScreenDC = GetDC(NULL); // Get the desktop DC
                    StretchBlt(hdc, 0, 0, width, height*ZOOM, hScreenDC, g_iSrcX, g_iSrcY, width/ZOOM, height, SRCCOPY);
                    ReleaseDC(NULL, hScreenDC);
                }

                #ifdef RENDER_CURSOR
                    // RenderScaledCursor(hdc, hWnd, width, height); // Render directly to hdc - this also works
                #endif
            }

            EndPaint(hWnd, &ps);
        }
        break;
//    case WM_LBUTTONDOWN: // Can't do anything on button down - because the input focus is on our window. Improve later.
    case WM_LBUTTONUP:
        {
            g_bTemporarilyHideMyWindow = false;

            HWND hFocusedChildWnd = g_hFocusedChildWnd; // To make sure it doesn't change mid-way

            if( hFocusedChildWnd )
            {
                POINT ptCurrent; GetCursorPos(&ptCurrent);

                POINT ptClient = ptCurrent; ScreenToClient(hWnd, &ptClient);

                POINT ptTarget = { ptClient.x / ZOOM + g_iSrcX, ptClient.y / ZOOM + g_iSrcY };

                POINT ptTargetAbs  = { (ptTarget.x  * 65536 + g_iScreenWidth/2) / g_iScreenWidth, (ptTarget.y  * 65536 + g_iScreenHeight/2) / g_iScreenHeight };
                POINT ptCurrentAbs = { (ptCurrent.x * 65536 + g_iScreenWidth/2) / g_iScreenWidth, (ptCurrent.y * 65536 + g_iScreenHeight/2) / g_iScreenHeight };

                g_iTakeCaretSnapshot = 10;// At least 4 is needed here... O.o
                g_bFreezeSrcXandSrcY = true;

                mouse_event(MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE, ptTargetAbs.x, ptTargetAbs.y, 0, 0);
                mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
                mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
                Sleep(1); // just in case...
                mouse_event(MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE, ptCurrentAbs.x, ptCurrentAbs.y, 0, 0);
            }
        }
        break;
    case WM_LBUTTONDOWN     :
    //case WM_LBUTTONUP     :
    case WM_LBUTTONDBLCLK   :
    case WM_RBUTTONDOWN     :
    case WM_RBUTTONUP       :
    case WM_RBUTTONDBLCLK   :
    case WM_MBUTTONDOWN     :
    case WM_MBUTTONUP       :
    case WM_MBUTTONDBLCLK   :
        g_bTemporarilyHideMyWindow = false;
        break;
    case WM_DESTROY:
        bDestroy = true;
        break;
    }

    // Execute the default handler
    lParam = DefWindowProc(hWnd, message, wParam, lParam);

    if (bDestroy)
    {
        #ifdef INVALIDATE_ON_TIMER
            KillTimer(hWnd, INVALIDATE_TIMER_ID); // Kill the timer
        #endif

        #ifdef INVALIDATE_ON_HOOK
            StopDesktopMonitor(); // Stop monitoring desktop for changes
        #endif

        PostQuitMessage(0);
    }

    return lParam;
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

void CALLBACK WinEventProc_ForFocusedClientWnd(HWINEVENTHOOK hWinEventHook, DWORD event, HWND hwnd,
                                              LONG idObject, LONG idChild, DWORD dwEventThread, DWORD dwmsEventTime)
{
    if( event == EVENT_OBJECT_FOCUS && idObject == OBJID_CLIENT )
    {
        if( g_hFocusedChildWnd != NULL)
        {
            // Check if this hwnd is a parent of g_hFocusedChildWnd
            static HWND hwndPrevMatchedParent = nullptr;
            static HWND hwndPrevMatchedParent_itsChild = nullptr;

            if( hwnd == hwndPrevMatchedParent )
            {
                g_hFocusedChildWnd = hwndPrevMatchedParent_itsChild;
            }
            else
            {
                BOOL bIsParent = FALSE;
                HWND hwndParent = g_hFocusedChildWnd;

                while( hwndParent = GetParent(hwndParent) )
                if( hwnd == hwndParent )
                {
                    bIsParent = TRUE;
                    hwndPrevMatchedParent = hwndParent;
                    hwndPrevMatchedParent_itsChild = g_hFocusedChildWnd;
                    break;
                }

                if( bIsParent == FALSE )
                    g_hFocusedChildWnd = hwnd;
            }
        }
        else
            g_hFocusedChildWnd = hwnd;
    }
}

LRESULT CALLBACK LowLevelKeyboardProc(int nCode, WPARAM wParam, LPARAM lParam)
{
    if (nCode == HC_ACTION)
    {
        KBDLLHOOKSTRUCT *pKeyBoard = (KBDLLHOOKSTRUCT *)lParam;
        if (wParam == WM_KEYDOWN || wParam == WM_SYSKEYDOWN)
        {
            DWORD virtualKey    = pKeyBoard->vkCode;
            UINT  mappedChar    = MapVirtualKey(virtualKey, MAPVK_VK_TO_CHAR); // maps the virtual key to a character
            bool  bPrintableKey = mappedChar && std::isprint(mappedChar); // Checks if the mapped character is printable

            if( bPrintableKey || virtualKey == VK_DELETE || virtualKey == VK_BACK )
                g_bTemporarilyHideMyWindow = false;
        }
    }
    return CallNextHookEx(NULL, nCode, wParam, lParam);
}