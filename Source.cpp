#pragma comment(lib,"gdiplus")
#pragma comment(lib,"shlwapi")

#include <windows.h>
#include <gdiplus.h>
#include <shlwapi.h>
using namespace Gdiplus;

TCHAR szClassName[] = TEXT("Image Converter");
HINSTANCE hInstance;

int GetEncoderClsid(const WCHAR* format, CLSID* pClsid)
{
	UINT num = 0, size = 0;
	GetImageEncodersSize(&num, &size);
	if (!size)return -1;
	ImageCodecInfo* pImageCodecInfo = (ImageCodecInfo*)GlobalAlloc(GPTR, size);
	if (!pImageCodecInfo)return -1;
	GetImageEncoders(num, size, pImageCodecInfo);
	for (UINT j = 0; j<num; ++j)
	{
		if (wcscmp(pImageCodecInfo[j].MimeType, format) == 0)
		{
			*pClsid = pImageCodecInfo[j].Clsid;
			GlobalFree(pImageCodecInfo);
			return j;
		}
	}
	GlobalFree(pImageCodecInfo);
	return -1;
}

BOOL SaveBitmapAs(LPCWSTR pszFileType, LPCWSTR pszFileName, Image* pImage)
{
	CLSID clsid;
	if (GetEncoderClsid(pszFileType, &clsid) >= 0)
	{
		pImage->Save(pszFileName, &clsid, 0);
		return TRUE;
	}
	return FALSE;
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	static HWND hStatic, hCombo;
	switch (msg)
	{
	case WM_CREATE:
		hStatic = CreateWindow(TEXT("Static"), 0, WS_CHILD | WS_VISIBLE | SS_BITMAP, 0, 64, 0, 0, hWnd, 0, ((LPCREATESTRUCT)(lParam))->hInstance, 0);
		hCombo = CreateWindow(TEXT("COMBOBOX"), 0, WS_CHILD | WS_VISIBLE | WS_VSCROLL | WS_BORDER | CBS_DROPDOWNLIST, 0, 0, 256, 2048, hWnd, 0, ((LPCREATESTRUCT)(lParam))->hInstance, 0);
		SendMessage(hCombo, CB_ADDSTRING, 0, (LPARAM)TEXT(".png"));
		SendMessage(hCombo, CB_ADDSTRING, 0, (LPARAM)TEXT(".jpg"));
		SendMessage(hCombo, CB_ADDSTRING, 0, (LPARAM)TEXT(".gif"));
		SendMessage(hCombo, CB_ADDSTRING, 0, (LPARAM)TEXT(".tif"));
		SendMessage(hCombo, CB_ADDSTRING, 0, (LPARAM)TEXT(".bmp"));
		SendMessage(hCombo, CB_ADDSTRING, 0, (LPARAM)TEXT(".ico"));
		SendMessage(hCombo, CB_ADDSTRING, 0, (LPARAM)TEXT(".emf"));
		SendMessage(hCombo, CB_ADDSTRING, 0, (LPARAM)TEXT(".wmf"));
		SendMessage(hCombo, CB_SETCURSEL, 0, 0);
		DragAcceptFiles(hWnd, TRUE);
		break;
	case WM_DROPFILES:
	{
		HDROP hDrop = (HDROP)wParam;
		TCHAR szFileName[MAX_PATH];
		UINT iFile, nFiles;
		nFiles = DragQueryFile((HDROP)hDrop, 0xFFFFFFFF, NULL, 0);
		for (iFile = 0; iFile<nFiles; ++iFile)
		{
			DragQueryFile(hDrop, iFile, szFileName, sizeof(szFileName));
			Image *imgTemp = Gdiplus::Image::FromFile(szFileName);
			if (imgTemp)
			{
				Bitmap*pBitmap = static_cast<Bitmap*>(imgTemp->Clone());
				HBITMAP hBitmap = NULL;
				Status status = pBitmap->GetHBITMAP(Color(0, 0, 0), &hBitmap);
				if (status == Ok)
				{
					SendMessage(hStatic, STM_SETIMAGE, IMAGE_BITMAP, (LPARAM)hBitmap);
				}
				delete pBitmap;
				TCHAR pszBuf[5] = { 0 };
				const DWORD nIndex = SendMessage(hCombo, CB_GETCURSEL, 0, 0);
				SendMessage(hCombo, CB_GETLBTEXT, nIndex, (LPARAM)pszBuf);
				PathRenameExtension(szFileName, pszBuf);
				switch (nIndex)
				{
				case 0:SaveBitmapAs(L"image/png", szFileName, imgTemp); break;
				case 1:SaveBitmapAs(L"image/jpeg", szFileName, imgTemp); break;
				case 2:SaveBitmapAs(L"image/gif", szFileName, imgTemp); break;
				case 3:SaveBitmapAs(L"image/tiff", szFileName, imgTemp); break;
				case 4:SaveBitmapAs(L"image/bmp", szFileName, imgTemp); break;
				case 5:SaveBitmapAs(L"image/x-icon", szFileName, imgTemp); break;
				case 6:SaveBitmapAs(L"image/x-emf", szFileName, imgTemp); break;
				case 7:SaveBitmapAs(L"image/x-wmf", szFileName, imgTemp); break;
				}
				delete imgTemp;
			}
		}
		SendMessage(hStatic, STM_SETIMAGE, IMAGE_BITMAP, 0);
		DragFinish(hDrop);
	}
	break;
	case WM_DESTROY:
		PostQuitMessage(0);
		break;
	default:
		return DefWindowProc(hWnd, msg, wParam, lParam);
	}
	return 0;
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPreInst, LPSTR pCmdLine, int nCmdShow)
{
	GdiplusStartupInput gdiplusStartupInput;
	ULONG_PTR gdiplusToken;
	GdiplusStartup(&gdiplusToken, &gdiplusStartupInput, 0);
	MSG msg;
	WNDCLASS wndclass = {
		CS_HREDRAW | CS_VREDRAW,
		WndProc,
		0,
		0,
		hInstance,
		0,
		LoadCursor(0,IDC_ARROW),
		(HBRUSH)(COLOR_WINDOW + 1),
		0,
		szClassName
	};
	RegisterClass(&wndclass);
	HWND hWnd = CreateWindow(
		szClassName,
		TEXT("Image Converter (変換する形式を選択して、画像ファイルをドラッグ＆ドロップしてください)"),
		WS_OVERLAPPEDWINDOW,
		CW_USEDEFAULT,
		0,
		CW_USEDEFAULT,
		0,
		0,
		0,
		hInstance,
		0
	);
	ShowWindow(hWnd, SW_SHOWDEFAULT);
	UpdateWindow(hWnd);
	while (GetMessage(&msg, 0, 0, 0))
	{
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}
	GdiplusShutdown(gdiplusToken);
	return (int)msg.wParam;
}
