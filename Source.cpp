#pragma comment(linker,"\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

#pragma comment(lib,"crypt32")
#pragma comment(lib,"wininet")

#include <windows.h>
#include <wininet.h>

#define API_KEY L"Input your API KEY"

TCHAR szClassName[] = TEXT("Window");

LPBYTE SendGoogleCloudVision(HWND hWnd, LPBYTE lpByte64)
{
	HINTERNET hSession = InternetOpen(TEXT("DOWNLOAD"), INTERNET_OPEN_TYPE_PRECONFIG, 0, 0, 0);
	if (!hSession)
	{
		MessageBox(hWnd, TEXT("WININETの初期化に失敗しました"), 0, MB_ICONERROR);
		return 0;
	}
	HINTERNET hConnection = InternetConnect(hSession, L"vision.googleapis.com", INTERNET_DEFAULT_HTTPS_PORT, 0, 0, INTERNET_SERVICE_HTTP, 0, 0);
	if (!hConnection)
	{
		InternetCloseHandle(hSession);
		MessageBox(hWnd, TEXT("サーバーへの接続に失敗しました"), 0, MB_ICONERROR);
		return 0;
	}
	WCHAR szPath[1024];
	wsprintf(szPath, L"v1/images:annotate?key=%s", API_KEY);
	HINTERNET hRequest = HttpOpenRequest(hConnection, TEXT("POST"), szPath, 0, 0, 0, INTERNET_FLAG_SECURE, 0);
	if (!hRequest)
	{
		InternetCloseHandle(hConnection);
		InternetCloseHandle(hSession);
		MessageBox(hWnd, TEXT("HTTPリクエストの作成に失敗しました"), 0, MB_ICONERROR);
		return 0;
	}
	wchar_t header[] = L"Content-Type: application/json; charset=utf-8";
	LPSTR data = (LPSTR)GlobalAlloc(0, 1024 + GlobalSize(lpByte64));
	lstrcpyA(data, "{\"requests\":[{\"features\":{\"type\":\"DOCUMENT_TEXT_DETECTION\",\"maxResults\":1},\"imageContext\":{\"languageHints\":\"ja\"},\"image\":{\"content\":\"");
	lstrcatA(data, (LPCSTR)lpByte64);
	lstrcatA(data, "\"}}]}");
	HttpSendRequest(hRequest, header, (DWORD)wcslen(header), (LPVOID)(data), (DWORD)strlen(data));
	GlobalFree(data);
	DWORD dwSize = 0;
	LPBYTE lpByte = (LPBYTE)GlobalAlloc(GMEM_FIXED, 1);
	DWORD dwRead;
	static BYTE szBuf[1024 * 4];
	LPBYTE lpTmp;
	for (;;)
	{
		if (!InternetReadFile(hRequest, szBuf, (DWORD)sizeof(szBuf), &dwRead) || !dwRead) break;
		lpTmp = (LPBYTE)GlobalReAlloc(lpByte, (SIZE_T)(dwSize + dwRead), GMEM_MOVEABLE);
		if (lpTmp == NULL) break;
		lpByte = lpTmp;
		CopyMemory(lpByte + dwSize, szBuf, dwRead);
		dwSize += dwRead;
	}
	lpByte[dwSize] = 0;
	InternetCloseHandle(hRequest);
	InternetCloseHandle(hConnection);
	InternetCloseHandle(hSession);
	return lpByte;
}

LPBYTE base64encode(LPCBYTE lpData, DWORD dwSize)
{
	DWORD dwResult = 0;
	if (CryptBinaryToStringA(lpData, dwSize, CRYPT_STRING_BASE64, 0, &dwResult)) {
		LPBYTE lpszBase64 = (LPBYTE)GlobalAlloc(GMEM_FIXED, dwResult + 1);
		if (CryptBinaryToStringA(lpData, dwSize, CRYPT_STRING_BASE64, (LPSTR)lpszBase64, &dwResult)) {
			return lpszBase64;
		}
	}
	return 0;
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	static HWND hEdit;
	static HFONT hFont;
	switch (msg)
	{
	case WM_CREATE:
		hEdit = CreateWindowEx(WS_EX_CLIENTEDGE, TEXT("EDIT"), 0, WS_VISIBLE | WS_CHILD | WS_VSCROLL | WS_HSCROLL | ES_MULTILINE | ES_AUTOHSCROLL | ES_AUTOVSCROLL, 0, 0, 0, 0, hWnd, 0, ((LPCREATESTRUCT)lParam)->hInstance, 0);
		SendMessage(hEdit, EM_LIMITTEXT, 0, 0);
		hFont = CreateFontW(-20, 0, 0, 0, FW_NORMAL, 0, 0, 0, SHIFTJIS_CHARSET, 0, 0, 0, 0, L"MS Shell Dlg");
		SendMessage(hEdit, WM_SETFONT, (WPARAM)hFont, 0);
		DragAcceptFiles(hWnd, TRUE);
		break;
	case WM_SIZE:
		MoveWindow(hEdit, 0, 0, LOWORD(lParam), HIWORD(lParam), TRUE);
		break;
	case WM_DROPFILES:
		{
			const UINT nFileCount = DragQueryFile((HDROP)wParam, 0xFFFFFFFF, NULL, 0);
			for (;;) {
				SetWindowText(hEdit, 0);
				WCHAR szFilePath[MAX_PATH];
				DragQueryFile((HDROP)wParam, 0, szFilePath, _countof(szFilePath));
				HANDLE hFile = CreateFile(szFilePath, GENERIC_READ, 0, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0);
				if (hFile == INVALID_HANDLE_VALUE) {
					MessageBox(hWnd, TEXT("ファイルが開けません"), 0, MB_OK);
					break;
				}
				else {
					LPBYTE lpByteFile = (LPBYTE)GlobalAlloc(GPTR, GetFileSize(hFile, 0) + 1);
					DWORD dwReadSize;
					ReadFile(hFile, lpByteFile, GetFileSize(hFile, 0), &dwReadSize, 0);
					LPBYTE lpByte64 = base64encode(lpByteFile, dwReadSize);
					GlobalFree(lpByteFile);
					CloseHandle(hFile);
					LPBYTE lpByte = SendGoogleCloudVision(hWnd, lpByte64);
					DWORD nLength = MultiByteToWideChar(CP_UTF8, 0, (LPCCH)lpByte, -1, 0, 0);
					LPWSTR lpszStringW = (LPWSTR)GlobalAlloc(0, nLength * sizeof(WCHAR));
					MultiByteToWideChar(CP_UTF8, 0, (LPCCH)lpByte, -1, lpszStringW, nLength);
					LPWSTR context = 0;
					wchar_t* p = wcstok_s(lpszStringW, L"\n", &context);
					SendMessage(hEdit, WM_SETREDRAW, FALSE, 0);
					while (p) {
						SendMessageW(hEdit, EM_REPLACESEL, 0, (LPARAM)p);
						SendMessageW(hEdit, EM_REPLACESEL, 0, (LPARAM)L"\r\n");
						p = wcstok_s(0, L"\n", &context);
					}
					SendMessage(hEdit, EM_SETSEL, 0, -1);
					SendMessage(hEdit, WM_SETREDRAW, TRUE, 0);
					GlobalFree(lpByte64);
					GlobalFree(lpszStringW);
					GlobalFree(lpByte);
					SetFocus(hEdit);
				}
				break;
			}
			DragFinish((HDROP)wParam);
		}
		break;
	case WM_DESTROY:
		DeleteObject(hFont);
		PostQuitMessage(0);
		break;
	default:
		return DefWindowProc(hWnd, msg, wParam, lParam);
	}
	return 0;
}

int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE hPreInst, LPWSTR pCmdLine, int nCmdShow)
{
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
		TEXT("Google Cloud Vision OCR をドラッグ＆ドロップされたテキスト情報を取得"),
		WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN,
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
	return (int)msg.wParam;
}
