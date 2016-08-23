////////////////////////////////////////////////////////////////////////////////
// Microsoft SQL Server Compact Sample Code
//
// Microsoft Confidential
//
// Copyright 1999 - 2002 Microsoft Corporation.  All Rights Reserved.
//
// File: NorthwindOleDb.cpp
//
// Comments:		Defines the entry point for the application.
//
// Notes: 
//					This example demonstrates the following functions
//      			1. Create Northwind sample database
//		        	2. Create Employees table
//			        3. Open a connection to Northwind database
//			        4. Insert employee sample data using OLE DB API
//			        5. Update employee info using OLE DB API
//			        6. Queries through ICommandText
//			        7. IRowsetIndex seek
//			        8. Insert BLOB to database using ISequentialStream 
//			        9. Load BLOB from database using ILockBytes
//			        10. Wrap employee data insertions in a transaction
//
////////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include <commctrl.h>
#include <aygshell.h>
#include <sipapi.h>
#include "Common.h"
#include "Employees.h"

// Global Variables:
//
HINSTANCE				g_hInst;				// The current instance
HWND					g_hwndCB;				// The command bar handle
Employees*				g_pEmployees;			// The pointer to employees object

static SHACTIVATEINFO	s_sai;

ATOM					MyRegisterClass	(HINSTANCE, LPTSTR);
BOOL					InitInstance	(HINSTANCE, int);
LRESULT CALLBACK		WndProc			(HWND, UINT, WPARAM, LPARAM);
HWND					CreateRpCommandBar(HWND);

int WINAPI WinMain(	HINSTANCE hInstance,
					HINSTANCE hPrevInstance,
					LPTSTR    lpCmdLine,
					int       nCmdShow)
{
	MSG msg;
	HACCEL hAccelTable;

	// Perform application initialization:
	if (!InitInstance (hInstance, nCmdShow)) 
	{
		return FALSE;
	}

	hAccelTable = LoadAccelerators(hInstance, (LPCTSTR)IDC_NORTHWINDOLEDB);

	// Main message loop:
	while (GetMessage(&msg, NULL, 0, 0)) 
	{
		if (!TranslateAccelerator(msg.hwnd, hAccelTable, &msg)) 
		{
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
	}

	return msg.wParam;
}

////////////////////////////////////////////////////////////////////////////////
//  FUNCTION: MyRegisterClass()
//
//  PURPOSE: Registers the window class.
//
//  COMMENTS:
//
//    It is important to call this function so that the application 
//    will get 'well formed' small icons associated with it.
//
////////////////////////////////////////////////////////////////////////////////
ATOM MyRegisterClass(HINSTANCE hInstance, LPTSTR szWindowClass)
{
	WNDCLASS	wc;

    wc.style			= CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc		= (WNDPROC) WndProc;
    wc.cbClsExtra		= 0;
    wc.cbWndExtra		= 0;
    wc.hInstance		= hInstance;
    wc.hIcon			= LoadIcon(hInstance, MAKEINTRESOURCE(IDI_NORTHWINDOLEDB));
    wc.hCursor			= 0;
    wc.hbrBackground	= (HBRUSH) GetStockObject(WHITE_BRUSH);
    wc.lpszMenuName		= 0;
    wc.lpszClassName	= szWindowClass;

	return RegisterClass(&wc);
}

////////////////////////////////////////////////////////////////////////////////
//  FUNCTION: InitInstance(HANDLE, int)
//
//  PURPOSE: Saves instance handle and creates main window
//
//  COMMENTS:
//
//    In this function, we save the instance handle in a global variable and
//    create and display the main program window.
//
////////////////////////////////////////////////////////////////////////////////
BOOL InitInstance(HINSTANCE hInstance, int nCmdShow)
{
	HWND	hWnd = NULL;
	TCHAR	szTitle[MAX_LOADSTRING];			// The title bar text
	TCHAR	szWindowClass[MAX_LOADSTRING];		// The window class name

	g_hInst = hInstance;		// Store instance handle in our global variable
	// Initialize global strings
	LoadString(hInstance, IDC_NORTHWINDOLEDB, szWindowClass, MAX_LOADSTRING);
	LoadString(hInstance, IDS_APP_TITLE, szTitle, MAX_LOADSTRING);

	//If it is already running, then focus on the window
	hWnd = FindWindow(szWindowClass, szTitle);	
	if (hWnd) 
	{
		// set focus to foremost child window
		// The "| 0x01" is used to bring any owned windows to the foreground and
		// activate them.
		SetForegroundWindow((HWND)((ULONG) hWnd | 0x00000001));
		return 0;
	} 

	MyRegisterClass(hInstance, szWindowClass);
	
	RECT	rect;
	GetClientRect(hWnd, &rect);
	
//	hWnd = CreateWindow(szWindowClass, szTitle, WS_VISIBLE | WS_NONAVDONEBUTTON,
	hWnd = CreateWindow(szWindowClass, szTitle, WS_VISIBLE,
		CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, NULL, NULL, hInstance, NULL);
	if (!hWnd)
	{	
		return FALSE;
	}
	//When the main window is created using CW_USEDEFAULT the height of the menubar (if one
	// is created is not taken into account). So we resize the window after creating it
	// if a menubar is present
	{
		RECT rc;
		GetWindowRect(hWnd, &rc);
		rc.bottom -= MENU_HEIGHT;
		if (g_hwndCB)
			MoveWindow(hWnd, rc.left, rc.top, rc.right, rc.bottom, FALSE);
	}

	ShowWindow(hWnd, nCmdShow);
	UpdateWindow(hWnd);

	return TRUE;
}

////////////////////////////////////////////////////////////////////////////////
//  FUNCTION: WndProc(HWND, unsigned, WORD, LONG)
//
//  PURPOSE:  Processes messages for the main window.
//
//  WM_COMMAND	- process the application menu
//  WM_PAINT	- Paint the main window
//  WM_DESTROY	- post a quit message and return
//
////////////////////////////////////////////////////////////////////////////////
LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	HDC hdc;
	int wmId, wmEvent;
	PAINTSTRUCT ps;
    BOOL bSuccess;

	switch (message) 
	{
		case WM_COMMAND:
			wmId    = LOWORD(wParam); 
			wmEvent = HIWORD(wParam); 
			// Parse the menu selections:
			switch (wmId)
			{	
				case IDOK:
					SendMessage(hWnd, WM_ACTIVATE, MAKEWPARAM(WA_INACTIVE, 0), (LPARAM)hWnd);
					SendMessage(hWnd, WM_CLOSE, 0, 0);
					break;
				case ID_FILE_EXIT:
					DestroyWindow(hWnd);
				default:
				   return DefWindowProc(hWnd, message, wParam, lParam);
			}
			break;
		case WM_CREATE:
			g_hwndCB = CreateRpCommandBar(hWnd);
            // Initialize the shell activate info structure
            memset (&s_sai, 0, sizeof (s_sai));
            s_sai.cbSize = sizeof (s_sai);

			// Create employee object
			//
            bSuccess = FALSE;
			g_pEmployees = new Employees(&bSuccess);

            if (!bSuccess || NULL == g_pEmployees)
            {
                delete g_pEmployees;
                g_pEmployees = NULL;
                DestroyWindow(hWnd);
                break;
            }

            // If failed to create employee dialog, exit
            //
			if (NULL == g_pEmployees->Create(hWnd, g_hInst))
            {
                DestroyWindow(hWnd);
            }

			break;
		case WM_PAINT:
			RECT rt;
			hdc = BeginPaint(hWnd, &ps);
			GetClientRect(hWnd, &rt);
			EndPaint(hWnd, &ps);
			break; 
		case WM_DESTROY:
			// Release employees object
			//
			delete g_pEmployees;

			CommandBar_Destroy(g_hwndCB);
			PostQuitMessage(0);
			break;
		case WM_ACTIVATE:
            // Notify shell of our activate message
			SHHandleWMActivate(hWnd, wParam, lParam, &s_sai, FALSE);
     		break;
		case WM_SETTINGCHANGE:
			SHHandleWMSettingChange(hWnd, wParam, lParam, &s_sai);
     		break;
		default:
			return DefWindowProc(hWnd, message, wParam, lParam);
   }
   return 0;
}

HWND CreateRpCommandBar(HWND hwnd)
{
	SHMENUBARINFO mbi;

	memset(&mbi, 0, sizeof(SHMENUBARINFO));
	mbi.cbSize     = sizeof(SHMENUBARINFO);
	mbi.hwndParent = hwnd;
	mbi.nToolBarId = IDM_MENU;
	mbi.hInstRes   = g_hInst;
	mbi.nBmpId     = 0;
	mbi.cBmpImages = 0;

	if (!SHCreateMenuBar(&mbi)) 
		return NULL;

	return mbi.hwndMB;
}

////////////////////////////////////////////////////////////////////////////////
// Function: EmployeesDlgProc
//
// Description: Handles messages for the employees dialog box
//
// Returns: The return value is the result of the message processing and depends 
// on the message sent
//
// Notes:
//
////////////////////////////////////////////////////////////////////////////////
LRESULT CALLBACK EmployeesDlgProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{  
	LRESULT	lResult = TRUE;

	switch(uMsg)
	{
		case WM_PAINT:
			if (g_pEmployees)
			{
				HDC				hDC;
				PAINTSTRUCT		ps;

				hDC = BeginPaint(hWnd, &ps);
				g_pEmployees->ShowEmployeePhoto();
				EndPaint(hWnd, &ps);
			}
			break;

		case WM_COMMAND:
			switch(LOWORD(wParam)) 
			{
				case IDC_COMBO_NAME: 
					if (HIWORD(wParam) == LBN_SELCHANGE) 
					{
						DWORD dwEmployeeID;
						DWORD dwCurSel;

						// Set current selection of employee name combobox to index 0,
						//
						dwCurSel = SendDlgItemMessage(hWnd, IDC_COMBO_NAME, CB_GETCURSEL, 0, 0);
						if (CB_ERR != dwCurSel)
						{
							HRESULT hr = NOERROR;

							// Retrieve current selected employee id from employee name combobox,
							// and update other employee info.
							//
							dwEmployeeID = SendDlgItemMessage(hWnd, IDC_COMBO_NAME, CB_GETITEMDATA, dwCurSel, 0);
							hr = g_pEmployees->LoadEmployeeInfo(dwEmployeeID);
							if (FAILED(hr))
							{
								MessageBox(NULL, L"Error - Update employee info", L"Northwind Oledb sample", MB_OK);
								break;
							}
						}
					}
					break;
				
				case IDC_BUTTON_SAVE:
					if (HIWORD(wParam) == BN_CLICKED) 
					{
						DWORD dwEmployeeID;
						DWORD dwCurSel;

						// Set current selection of employee name combobox to index 0,
						//
						dwCurSel = SendDlgItemMessage(hWnd, IDC_COMBO_NAME, CB_GETCURSEL, 0, 0);
						if (CB_ERR != dwCurSel)
						{
							HRESULT hr = NOERROR;

							// Retrieve current selected employee id from employee name combobox,
							// and save employee info to database.
							//
							dwEmployeeID = SendDlgItemMessage(hWnd, IDC_COMBO_NAME, CB_GETITEMDATA, dwCurSel, 0);
							hr = g_pEmployees->SaveEmployeeInfo(dwEmployeeID);
							if (FAILED(hr))
							{
								MessageBox(NULL, L"Error - Save employee info", L"Northwind Oledb sample", MB_OK);
								break;
							}
						}

						break;
					}
					break;

				case IDC_BUTTON_EXIT:
					if (HIWORD(wParam) == BN_CLICKED) 
					{
						// Release employees object
						//
						delete g_pEmployees;

						CommandBar_Destroy(g_hwndCB);
						PostQuitMessage(0);
						break;
					}
					break;

				default:
					break;
			}
			break;

		default:
			return DefWindowProc(hWnd, uMsg, wParam, lParam);
			break;
	}

	return (lResult);
}

