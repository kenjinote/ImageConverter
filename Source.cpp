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
	static HWND hStatic, hCombo1, hCombo2;
	switch (msg)
	{
	case WM_CREATE:
		hStatic = CreateWindow(TEXT("Static"), 0, WS_CHILD | WS_VISIBLE | SS_BITMAP, 0, 80, 0, 0, hWnd, 0, ((LPCREATESTRUCT)(lParam))->hInstance, 0);
		hCombo1 = CreateWindow(TEXT("COMBOBOX"), 0, WS_CHILD | WS_VISIBLE | WS_VSCROLL | WS_BORDER | CBS_DROPDOWNLIST, 0, 0, 256, 2048, hWnd, 0, ((LPCREATESTRUCT)(lParam))->hInstance, 0);
		SendMessage(hCombo1, CB_ADDSTRING, 0, (LPARAM)TEXT(".png"));
		SendMessage(hCombo1, CB_ADDSTRING, 0, (LPARAM)TEXT(".jpg"));
		SendMessage(hCombo1, CB_ADDSTRING, 0, (LPARAM)TEXT(".gif"));
		SendMessage(hCombo1, CB_ADDSTRING, 0, (LPARAM)TEXT(".tif"));
		SendMessage(hCombo1, CB_ADDSTRING, 0, (LPARAM)TEXT(".bmp"));
		SendMessage(hCombo1, CB_ADDSTRING, 0, (LPARAM)TEXT(".ico"));
		SendMessage(hCombo1, CB_ADDSTRING, 0, (LPARAM)TEXT(".emf"));
		SendMessage(hCombo1, CB_ADDSTRING, 0, (LPARAM)TEXT(".wmf"));
		SendMessage(hCombo1, CB_SETCURSEL, 0, 0);
		hCombo2 = CreateWindow(TEXT("COMBOBOX"), 0, WS_CHILD | WS_VISIBLE | WS_VSCROLL | WS_BORDER | CBS_DROPDOWNLIST, 0, 40, 256, 2048, hWnd, 0, ((LPCREATESTRUCT)(lParam))->hInstance, 0);
		SendMessage(hCombo2, CB_ADDSTRING, 0, (LPARAM)TEXT("PixelFormat1bppIndexed"));
		SendMessage(hCombo2, CB_ADDSTRING, 0, (LPARAM)TEXT("PixelFormat4bppIndexed"));
		SendMessage(hCombo2, CB_ADDSTRING, 0, (LPARAM)TEXT("PixelFormat8bppIndexed"));
		SendMessage(hCombo2, CB_ADDSTRING, 0, (LPARAM)TEXT("PixelFormat16bppGrayScale"));
		SendMessage(hCombo2, CB_ADDSTRING, 0, (LPARAM)TEXT("PixelFormat16bppRGB555"));
		SendMessage(hCombo2, CB_ADDSTRING, 0, (LPARAM)TEXT("PixelFormat16bppRGB565"));
		SendMessage(hCombo2, CB_ADDSTRING, 0, (LPARAM)TEXT("PixelFormat16bppARGB1555"));
		SendMessage(hCombo2, CB_ADDSTRING, 0, (LPARAM)TEXT("PixelFormat24bppRGB"));
		SendMessage(hCombo2, CB_ADDSTRING, 0, (LPARAM)TEXT("PixelFormat32bppRGB"));
		SendMessage(hCombo2, CB_ADDSTRING, 0, (LPARAM)TEXT("PixelFormat32bppARGB"));
		SendMessage(hCombo2, CB_ADDSTRING, 0, (LPARAM)TEXT("PixelFormat32bppPARGB"));
		SendMessage(hCombo2, CB_ADDSTRING, 0, (LPARAM)TEXT("PixelFormat48bppRGB"));
		SendMessage(hCombo2, CB_ADDSTRING, 0, (LPARAM)TEXT("PixelFormat64bppARGB"));
		SendMessage(hCombo2, CB_ADDSTRING, 0, (LPARAM)TEXT("PixelFormat64bppPARGB"));
		SendMessage(hCombo2, CB_ADDSTRING, 0, (LPARAM)TEXT("PixelFormat32bppCMYK"));
		SendMessage(hCombo2, CB_SETCURSEL, 7, 0);
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
			Gdiplus::Bitmap *imgTemp = Gdiplus::Bitmap::FromFile(szFileName);
			if (imgTemp)
			{
				Gdiplus::Bitmap*pBitmap = 0;
				switch (SendMessage(hCombo2, CB_GETCURSEL, 0, 0))
				{
				case 0:pBitmap = imgTemp->Clone(0, 0, imgTemp->GetWidth(), imgTemp->GetHeight(), PixelFormat1bppIndexed); break;
				case 1:pBitmap = imgTemp->Clone(0, 0, imgTemp->GetWidth(), imgTemp->GetHeight(), PixelFormat4bppIndexed); break;
				case 2:pBitmap = imgTemp->Clone(0, 0, imgTemp->GetWidth(), imgTemp->GetHeight(), PixelFormat8bppIndexed); break;
				case 3:pBitmap = imgTemp->Clone(0, 0, imgTemp->GetWidth(), imgTemp->GetHeight(), PixelFormat16bppGrayScale); break;
				case 4:pBitmap = imgTemp->Clone(0, 0, imgTemp->GetWidth(), imgTemp->GetHeight(), PixelFormat16bppRGB555); break;
				case 5:pBitmap = imgTemp->Clone(0, 0, imgTemp->GetWidth(), imgTemp->GetHeight(), PixelFormat16bppRGB565); break;
				case 6:pBitmap = imgTemp->Clone(0, 0, imgTemp->GetWidth(), imgTemp->GetHeight(), PixelFormat16bppARGB1555); break;
				case 7:pBitmap = imgTemp->Clone(0, 0, imgTemp->GetWidth(), imgTemp->GetHeight(), PixelFormat24bppRGB); break;
				case 8:pBitmap = imgTemp->Clone(0, 0, imgTemp->GetWidth(), imgTemp->GetHeight(), PixelFormat32bppRGB); break;
				case 9:pBitmap = imgTemp->Clone(0, 0, imgTemp->GetWidth(), imgTemp->GetHeight(), PixelFormat32bppARGB); break;
				case 10:pBitmap = imgTemp->Clone(0, 0, imgTemp->GetWidth(), imgTemp->GetHeight(), PixelFormat32bppPARGB); break;
				case 11:pBitmap = imgTemp->Clone(0, 0, imgTemp->GetWidth(), imgTemp->GetHeight(), PixelFormat48bppRGB); break;
				case 12:pBitmap = imgTemp->Clone(0, 0, imgTemp->GetWidth(), imgTemp->GetHeight(), PixelFormat64bppARGB); break;
				case 13:pBitmap = imgTemp->Clone(0, 0, imgTemp->GetWidth(), imgTemp->GetHeight(), PixelFormat64bppPARGB); break;
				case 14:pBitmap = imgTemp->Clone(0, 0, imgTemp->GetWidth(), imgTemp->GetHeight(), PixelFormat32bppCMYK); break;
				}
				if (pBitmap)
				{
					HBITMAP hBitmap = NULL;
					Status status = pBitmap->GetHBITMAP(Color(0, 0, 0), &hBitmap);
					if (status == Ok)
					{
						SendMessage(hStatic, STM_SETIMAGE, IMAGE_BITMAP, (LPARAM)hBitmap);
					}
					TCHAR pszBuf[5] = { 0 };
					const DWORD_PTR nIndex = SendMessage(hCombo1, CB_GETCURSEL, 0, 0);
					SendMessage(hCombo1, CB_GETLBTEXT, nIndex, (LPARAM)pszBuf);
					PathRenameExtension(szFileName, pszBuf);
					switch (nIndex)
					{
					case 0:SaveBitmapAs(L"image/png", szFileName, pBitmap); break;
					case 1:SaveBitmapAs(L"image/jpeg", szFileName, pBitmap); break;
					case 2:SaveBitmapAs(L"image/gif", szFileName, pBitmap); break;
					case 3:SaveBitmapAs(L"image/tiff", szFileName, pBitmap); break;
					case 4:SaveBitmapAs(L"image/bmp", szFileName, pBitmap); break;
					case 5:SaveBitmapAs(L"image/x-icon", szFileName, pBitmap); break;
					case 6:SaveBitmapAs(L"image/x-emf", szFileName, pBitmap); break;
					case 7:SaveBitmapAs(L"image/x-wmf", szFileName, pBitmap); break;
					}
					delete pBitmap;
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
