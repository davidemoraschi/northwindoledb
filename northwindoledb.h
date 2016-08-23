/***********************************************************************
Copyright (c) 1999 - 2002, Microsoft Corporation
All Rights Reserved.
***********************************************************************/

#if !defined(AFX_NORTHWINDOLEDB_H__1E5E7D23_3B1E_4767_B01E_434D7656C369__INCLUDED_)
#define AFX_NORTHWINDOLEDB_H__1E5E7D23_3B1E_4767_B01E_434D7656C369__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "employees.h"

// Forward declarations of functions included in this code module:
ATOM				MyRegisterClass	(HINSTANCE, LPTSTR);
BOOL				InitInstance	(HINSTANCE, int);
LRESULT CALLBACK	WndProc			(HWND, UINT, WPARAM, LPARAM);
LRESULT CALLBACK	About			(HWND, UINT, WPARAM, LPARAM);
LRESULT CALLBACK EmployeesDlgProc(HWND, UINT, WPARAM, LPARAM);
HWND				CreateRpCommandBar(HWND);

// Global Variables:
HINSTANCE			g_hInst;				// The current instance
HWND				g_hwndCB;				// The command bar handle
Employees*			g_pEmployees;

#endif // !defined(AFX_NORTHWINDOLEDB_H__1E5E7D23_3B1E_4767_B01E_434D7656C369__INCLUDED_)
