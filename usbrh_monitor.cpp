/**
 * @file
 * @brief Strawberry Linux USB温度・湿度計モジュール・キット(メーカー品番：USBRH)用のモニタリングツールです。
 * @author AKIYAMA Kouhei
 * @since 2010-07-11
 */
#include <winsock2.h>
#include <windows.h>
#include <tchar.h>
#include <vector>
#include <boost/shared_ptr.hpp>
#include <ctime>
#include <cstdio>
#include <string>

#include "resource.h"

/**
 * USBMeter.dllへのアクセスを提供する関数をまとめた名前空間です。
 * 一番最初に usbmeter::LoadUSBMeterDLL を呼び出してください。
 */
namespace usbmeter {
	void ERROR_WINAPI(TCHAR *msg)
	{
		MessageBox(NULL, msg, _T("Error"), MB_OK);
	}

	typedef BSTR (__stdcall *FPGetVers)(BSTR dev);
	typedef BSTR (__stdcall *FPFindUSB)(LONG *index);
	typedef LONG (__stdcall *FPGetTempHumid)(BSTR dev, double *temp, double *humid);
	typedef LONG (__stdcall *FPControlIO)(BSTR dev, LONG port, LONG val);
	typedef LONG (__stdcall *FPSetHeater)(BSTR dev, LONG val);
	typedef LONG (__stdcall *FPGetTempHumidTrue)(BSTR dev, double *temp, double *humid);

	static FPGetVers fpGetVers;
	static FPFindUSB fpFindUSB;
	static FPGetTempHumid fpGetTempHumid;
	static FPControlIO fpControlIO;
	static FPSetHeater fpSetHeater;
	static FPGetTempHumidTrue fpGetTempHumidTrue;

	BSTR GetVers(BSTR dev) { return (*fpGetVers)(dev);}
	BSTR FindUSB(LONG *index) { return (*fpFindUSB)(index);}
	LONG GetTempHumid(BSTR dev, double *temp, double *humid) { return (*fpGetTempHumid)(dev, temp, humid);}
	LONG ControlIO(BSTR dev, LONG port, LONG val) { return (*fpControlIO)(dev, port, val);}
	LONG SetHeater(BSTR dev, LONG val) { return (*fpSetHeater)(dev, val);}
	LONG GetTempHumidTrue(BSTR dev, double *temp, double *humid) { return (*fpGetTempHumidTrue)(dev, temp, humid);}

	template<typename FP>
	bool GetFunction(HMODULE hModule, LPCSTR fname, FP *fpp)
	{
		FP fp = reinterpret_cast<FP>(GetProcAddress(hModule, fname));
		if(!fp){
			ERROR_WINAPI(_T("GetProcAddress Failed."));
			return false;
		}
		*fpp = fp;
		return true;
	}

	bool LoadUSBMeterDLL(void)
	{
		HMODULE hModule = LoadLibrary(_T("USBMeter.dll"));
		if(!hModule){
			ERROR_WINAPI(_T("Failed to load USBMeter.dll. You can download USBMeter.dll from http://strawberry-linux.com/download/ (#52002)."));
			return false;
		}
		return GetFunction(hModule, "_GetVers@4", &fpGetVers)
			&& GetFunction(hModule, "_FindUSB@4", &fpFindUSB)
			&& GetFunction(hModule, "_GetTempHumid@12", &fpGetTempHumid)
			&& GetFunction(hModule, "_ControlIO@12", &fpControlIO)
			&& GetFunction(hModule, "_SetHeater@8", &fpSetHeater)
			&& GetFunction(hModule, "_GetTempHumidTrue@12", &fpGetTempHumidTrue);
	}
};

/**
 * BSTRを解放します。SysFreeString(p)を実行します。shared_ptr用の削除子として使います。
 */
void ReleaseBSTR(BSTR p) { ::SysFreeString(p); }


/**
 * 接続しているusbrhデバイスのデバイス名を列挙します。
 */
void EnumerateDevices(std::vector<boost::shared_ptr<OLECHAR> > *devices)
{
	LONG index = 0;
	for(;;){
		BSTR dev = usbmeter::FindUSB(&index); //UnicodeビルドでもANSIで返ってくるので注意。
		// CSimpleArray<CAdapt<CComBSTR> >なんかを使うべきなのかもしれないけれどATLには詳しくないしコピーとかが気になるので、とりあえずstd::vectorとshared_ptrを使うことにする。
		// std::vector<BSTR>で解放処理は行わないってのもありかも。どうせデバイスを列挙するのは一回だけだし。
		boost::shared_ptr<OLECHAR> devptr(dev, ReleaseBSTR);
		if(::SysStringLen(dev) == 0){
			break;
		}
		devices->push_back(devptr);
	}
	return;
}









// --------------------------------------------------------------------------
// メインダイアログボックス
// --------------------------------------------------------------------------

/**
 * アプリケーションの中心となるダイアログボックスウィンドウです。
 *
 * ダイアログには現在の計測データや設定を表示します。
 * タイマーイベントで計測やログ出力を行います。
 * 起動時はタスクアイコンを登録して非表示状態になります。
 */
class MainDialogBox
{
	static const UINT WM_NOTIFY_ICON = (WM_USER + 100);///< システムトレイアイコンからの通知用。
	static const UINT TIMER_ID = 1; ///< モニタ更新、経過時間確認に使用するタイマーのID。
	static const UINT TIMER_PERIOD = 1000; ///< モニタ更新頻度、経過時間確認頻度。

	const HINSTANCE hInst_;
	const std::vector<boost::shared_ptr<OLECHAR> > &devices_;

	HWND hWnd_;
	HMENU hPopupMenu_;
	NOTIFYICONDATA NotifyIconData_;
	HICON hIconApp_;

	SYSTEMTIME prevTime_;

	bool visible_;

	typedef std::basic_string<TCHAR> StdString;
	StdString outputFilename_;  //example: temp_humid.csv
	StdString outputHTTPURL_;   //example: example.jp/temp_humid/record.pl
public:
	MainDialogBox(HINSTANCE hInst, const std::vector<boost::shared_ptr<OLECHAR> > &devices)
		: hInst_(hInst), devices_(devices), hWnd_(NULL), hPopupMenu_(), hIconApp_(), visible_()
		, outputFilename_()
		, outputHTTPURL_()
	{}
	virtual ~MainDialogBox(){}

	HWND Create(void)
	{
		if(hWnd_){
			return NULL;
		}
		hWnd_ = ::CreateDialogParam(
			hInst_,
			MAKEINTRESOURCE(IDD_MAIN),
			NULL,
			(DLGPROC)DlgProcStatic,
			(LPARAM)this);
		return hWnd_;
	}


private:
	static INT_PTR CALLBACK DlgProcStatic(
		HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
	{
		if(msg == WM_INITDIALOG){
			::SetWindowLong(hWnd, DWL_USER, lParam);
		}
		MainDialogBox *pThis;
		pThis = (MainDialogBox *)::GetWindowLong(hWnd, DWL_USER);

		if(pThis){
			return pThis->DlgProc(hWnd, msg, wParam, lParam);
		}
		else{
			return FALSE;
		}
	}

	/**
	 * ダイアログメッセージをメッセージを各関数に振り分けるだけです。
	 * @return TRUE(1)のときメッセージを処理
	 */
	INT_PTR DlgProc(
		HWND hWnd, UINT Message, WPARAM wParam, LPARAM lParam)
	{
		switch(Message){
		case WM_INITDIALOG:
			return OnInitDialog(hWnd, (HWND)wParam, lParam);
		case WM_DESTROY:
			return OnDestroy(hWnd);
		case WM_CLOSE:
			return OnClose(hWnd);
		case WM_COMMAND:
			return OnCommand(
				hWnd,
				(int)(LOWORD(wParam)),
				(HWND)(lParam),
				(UINT)HIWORD(wParam));
		case WM_NOTIFY_ICON:
			OnNotifyIcon(hWnd, wParam, lParam);
			return TRUE;
		case WM_TIMER:
			OnTimer(hWnd, wParam);
			return 0;
		}
		return FALSE;
	}

	/**
	 * WM_INITDIALOGメッセージの処理
	 */
	BOOL OnInitDialog(HWND hWnd, HWND hWndFocus, LPARAM lParam)
	{
		// 設定の読み込み
		LoadSettingsFromRegistory();
		CopySettingsVariablesToDlgItems();

		// ポップアップメニュー作成
		hPopupMenu_ = ::LoadMenu(hInst_, _T("POPUPMENU"));

		// アイコン読み込み
		hIconApp_ = ::LoadIcon(hInst_, _T("IDI_APP"));

		// アイコン設定
		::ZeroMemory(&NotifyIconData_, sizeof(NotifyIconData_));
		NotifyIconData_.cbSize = sizeof(NotifyIconData_);
		NotifyIconData_.hWnd   = hWnd;
		NotifyIconData_.uID    = NULL;
		NotifyIconData_.uFlags = NIF_MESSAGE | NIF_ICON | NIF_TIP;
		NotifyIconData_.uCallbackMessage = WM_NOTIFY_ICON;
		NotifyIconData_.hIcon  = hIconApp_;
		::lstrcpy(NotifyIconData_.szTip, _T("USBRH Thermo-Hygrometer"));

		::Shell_NotifyIcon(NIM_ADD, &NotifyIconData_);

		// タイマー設定
		::GetLocalTime(&prevTime_);
		::SetTimer(hWnd, TIMER_ID, TIMER_PERIOD, NULL);

		return TRUE;
	}

	/**
	 * WM_DESTROYメッセージ処理
	 */
	int OnDestroy(HWND hWnd)
	{
		PostQuitMessage(0);

		// ポップアップメニューを削除
		DestroyMenu(hPopupMenu_);

		// インジゲータ領域から削除
		Shell_NotifyIcon(NIM_DELETE, &NotifyIconData_);

		// タイマーを削除
		KillTimer(hWnd, TIMER_ID);

		return TRUE;
	}

	/**
	 * ダイアログボックスを表示状態にします。
	 */
	void Show(void)
	{
		if(!visible_){
			visible_ = true;
			ShowWindow(hWnd_, SW_SHOW);
		}
	}

	/**
	 * ダイアログボックスを非表示状態にします。
	 */
	void Hide(void)
	{
		if(visible_){
			visible_ = false;
			ShowWindow(hWnd_, SW_HIDE);
		}
	}

	/**
	 * WM_CLOSEメッセージ処理
	 */
	int OnClose(HWND hWnd)
	{
		Hide();
		return TRUE;
	}

	/**
	 * WM_COMMANDメッセージ処理
	 * @param hWnd ウィンドウハンドル	
	 * @param id 項目ID、コントロールID、アクセラレータID
	 * @param hWndCtl コントロールのハンドル
	 * @param codeNotify 通知コード
	 */
	int OnCommand(HWND hWnd,
		int id,
		HWND hwndctl,
		UINT codeNotify)
	{
		switch (id) {
		case IDM_MONITOR:
			Show();
			break;
		case IDM_ABOUT:
			MessageBox(
				hWnd,
				_T("USBRH Monitor\n2010 AKIYAMA Kouhei"),
				_T("タイトル"),
				MB_OK);
			break;
		case IDM_EXIT:
			DestroyWindow(hWnd);
			break;
		// メインダイアログからのコマンド
		case IDOK:
			CopySettingsDlgItemsToVariables();
			StoreSettingsToRegistory();
			OnClose(hWnd);
			return TRUE;
		case IDCANCEL:
			CopySettingsVariablesToDlgItems();
			return TRUE;
		}

		return TRUE;
	}


	/**
	 * システムトレイメッセージ
	 * WM_NOTIFY_ICONメッセージ処理
	 * @param hWnd ウィンドウハンドル
	 * @param wParam
	 * @param lParam メッセージ種類
	 */
	void OnNotifyIcon(HWND hWnd, WPARAM wParam, LPARAM lParam)
	{
		switch(lParam){
		case WM_LBUTTONDOWN:
		case WM_RBUTTONDOWN:
			POINT pt;
			GetCursorPos(&pt);

			HMENU hMenu = GetSubMenu(hPopupMenu_, 0);
			SetForegroundWindow( hWnd ) ;

			TrackPopupMenu(hMenu
				, TPM_LEFTBUTTON | TPM_BOTTOMALIGN | TPM_RIGHTALIGN
				, pt.x, pt.y, 0, hWnd, NULL);

			break;
		}
	}


	// 設定の保存と読み込み
	static const TCHAR *GetAppRegistryKey(void)
	{
		return _T("SOFTWARE\\Misohena Laboratories\\usbrh_monitor");
	}

	void LoadSettingsFromRegistory(void)
	{
		struct Reg {
			static StdString ReadString(HKEY hKey, const TCHAR *name, const TCHAR *defaultValue = _T(""))
			{
				const UINT MAX_STRING_LENGTH = 1024;
				TCHAR value[MAX_STRING_LENGTH + 1];
				ZeroMemory(value, sizeof(value));
				DWORD type = 0;
				DWORD bytes = MAX_STRING_LENGTH * sizeof(TCHAR);
				LONG result = RegQueryValueEx(hKey, name, NULL, &type, (BYTE *)value, &bytes);
				if(result == ERROR_SUCCESS && type == REG_SZ){
					return StdString(value);
				}
				return defaultValue;
			}
		};
		HKEY hKey;
		if(::RegOpenKeyEx(HKEY_CURRENT_USER, GetAppRegistryKey(), NULL, KEY_ALL_ACCESS, &hKey) != ERROR_SUCCESS){
			return; //error
		}
		outputFilename_ = Reg::ReadString(hKey, "OutputFilename");
		outputHTTPURL_ = Reg::ReadString(hKey, "OutputHTTPURL");
		RegCloseKey(hKey);
	}

	void StoreSettingsToRegistory(void)
	{
		struct Reg {
			static void WriteString(HKEY hKey, const TCHAR *name, const StdString &value)
			{
				const DWORD bytes = (DWORD)((value.size() + 1) * sizeof(TCHAR));
				LONG result = RegSetValueEx(hKey, name, NULL, REG_SZ, (BYTE *)value.c_str(), bytes);
				//if(result == ERROR_SUCCESS){
			}
		};
		HKEY hKey;
		DWORD disposition;
		if(::RegCreateKeyEx(HKEY_CURRENT_USER, GetAppRegistryKey(), NULL, _T(""), REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &hKey, &disposition) != ERROR_SUCCESS){
			return; // error
		}
		Reg::WriteString(hKey, "OutputFilename", outputFilename_);
		Reg::WriteString(hKey, "OutputHTTPURL", outputHTTPURL_);
		RegCloseKey(hKey);
	}

	/**
	 * 設定をメンバ変数からダイアログへコピーします。
	 * レジストリから読み込んだ後やキャンセルの時などに使用します。
	 */
	void CopySettingsVariablesToDlgItems(void)
	{
		SetDlgItemText(hWnd_, IDC_EDIT_OUTPUT_FILE, outputFilename_.c_str());
		SetDlgItemText(hWnd_, IDC_EDIT_OUTPUT_URL, outputHTTPURL_.c_str());
	}

	/**
	 * 設定をダイアログからメンバ変数へコピーします。
	 * OKボタンを押したときに使用します。
	 */
	void CopySettingsDlgItemsToVariables(void)
	{
		UINT len;
		TCHAR str[MAX_PATH];
		len = GetDlgItemText(hWnd_, IDC_EDIT_OUTPUT_FILE, str, sizeof(str) - 1);
		outputFilename_ = StdString(str, len);

		len = GetDlgItemText(hWnd_, IDC_EDIT_OUTPUT_URL, str, sizeof(str) - 1);
		outputHTTPURL_ = StdString(str, len);
	}


	// タイマーによる計測と出力


	void OnTimer(HWND hWnd, UINT TimerId)
	{
		SYSTEMTIME currTime;
		::GetLocalTime(&currTime);
		SYSTEMTIME prevTime = prevTime_;
		prevTime_ = currTime;

		if(currTime.wMinute != prevTime.wMinute){ ///@todo 現在は一分間隔固定になっているので、可変にしたい。
			Log(currTime, &MainDialogBox::OutputToFile);
			Log(currTime, &MainDialogBox::OutputToHTTP);
		}

		if(visible_){ // ダイアログを表示しているときだけ
			Log(currTime, &MainDialogBox::OutputToDialog);
		}
	}

	void Log(const SYSTEMTIME &currTime, void (MainDialogBox::*func)(const SYSTEMTIME &, double, double))
	{
		if(devices_.empty()){
			return;
		}

		///@todo 複数のデバイスに対応する。usbrhを一つしか持っていないので対応する気にならない。どういう出力にすべきかもよく分からない。

		double temp = 0;
		double humid = 0;
		if(usbmeter::GetTempHumidTrue(devices_[0].get(), &temp, &humid) != 0){
			return;
		}

		(this->*func)(currTime, temp, humid);
	}

	/**
	 * 計測データをファイルへ追記出力します。
	 */
	void OutputToFile(const SYSTEMTIME &currTime, double temp, double humid)
	{
		if(outputFilename_.empty()){
			return;
		}
		char str[512];
		const int size = std::sprintf(
			str, "%04d/%02d/%02d %02d:%02d:%02d,%lf,%lf\r\n",
			currTime.wYear,
			currTime.wMonth,
			currTime.wDay,
			currTime.wHour,
			currTime.wMinute,
			currTime.wSecond,
			temp,
			humid);
		HANDLE hFile = ::CreateFile(outputFilename_.c_str(), GENERIC_WRITE, 0, NULL, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
		::SetFilePointer(hFile, 0, NULL, FILE_END);
		DWORD written = 0;
		::WriteFile(hFile, str, size, &written, NULL);
		::CloseHandle(hFile);
	}

	/**
	 * 計測データをダイアログボックスへ出力します。
	 */
	void OutputToDialog(const SYSTEMTIME &currTime, double temp, double humid)
	{
		TCHAR str[1024];
		std::sprintf(
			str, "%s\n%04d/%02d/%02d %02d:%02d:%02d\ntemp=%lf humid=%lf",
			devices_[0].get(),
			currTime.wYear,
			currTime.wMonth,
			currTime.wDay,
			currTime.wHour,
			currTime.wMinute,
			currTime.wSecond,
			temp,
			humid);

		SetDlgItemText(hWnd_, IDC_STATIC_STATUS, str);
	}

	/**
	 * 計測データをHTTPのGETメソッドのパラメータとして出力します。
	 * http://example.jp/record.cgi?temp=25.543&humid=45.123 のような形式でアクセスします。
	 */
	void OutputToHTTP(const SYSTEMTIME &currTime, double temp, double humid)
	{
		if(outputHTTPURL_.empty()){
			return;
		}

		///@todo ポート指定できるようにする。
		///@todo 先頭に http:// と入れても大丈夫なようにする。
		const StdString::size_type firstSlashPos = outputHTTPURL_.find_first_of('/');
		const StdString destHostName = (firstSlashPos != StdString::npos) ? outputHTTPURL_.substr(0, firstSlashPos) : outputHTTPURL_;
		const StdString destURL      = (firstSlashPos != StdString::npos) ? outputHTTPURL_.substr(firstSlashPos) : "/";

		struct Local
		{
			static bool connectToHost(SOCKET &sock, const char *hostname)
			{
				sockaddr_in server;
				server.sin_family = AF_INET;
				server.sin_port = htons(80);
				server.sin_addr.S_un.S_addr = inet_addr(hostname);

				if(server.sin_addr.S_un.S_addr == 0xffffffff) {
					hostent *host = gethostbyname(hostname);
					if(host == NULL){
						return false;
					}

					unsigned int **addrptr = (unsigned int **)host->h_addr_list;
					for(; *addrptr != NULL; ++addrptr){
						server.sin_addr.S_un.S_addr = *(*addrptr);

						if(connect(sock, (sockaddr *)&server, sizeof(server)) == 0){
							break;
						}
					}
					if(*addrptr == NULL){
						return false;
					}
				}
				else{
					if (connect(sock, (sockaddr *)&server, sizeof(server)) != 0){
						return false;
					}
				}
				return true;
			}
		};

		SOCKET sock;
		sock = socket(AF_INET, SOCK_STREAM, 0);
		if(sock != INVALID_SOCKET) {

			// 接続
			if(Local::connectToHost(sock, destHostName.c_str())){

				// リクエスト作成
				char requestParams[1024];
				std::sprintf(requestParams, "?temp=%lf&humid=%lf", temp, humid);
				std::string request = std::string("GET ") + destURL + requestParams + " HTTP/1.0\r\n\r\n";

				// HTTPリクエスト送信
				bool error = false;
				int n = send(sock, request.c_str(), request.size(), 0);
				if(n < 0) {
					error = true;
				}

				// HTTPメッセージ受信
				std::string result;
				while(n > 0){
					char buf[128];
					n = recv(sock, buf, sizeof(buf), 0);
					if(n < 0) {
						error = true;
						break;
					}

					// 受信結果
					result.append(buf, n);
				}
			}

			closesocket(sock);
		}
	}

};// class MainDialogBox


int WINAPI WinMain(HINSTANCE hInst, HINSTANCE hPrevInst, LPSTR lpszCmdParam, int nCmdShow)
{
	// 温度計・湿度計デバイスの初期化
	if(!usbmeter::LoadUSBMeterDLL()){
		return -1;
	}
	std::vector<boost::shared_ptr<OLECHAR> > devices;
	EnumerateDevices(&devices);
	if(devices.empty()){
		MessageBox(NULL, _T("No USB meter device found."), _T("Error"), MB_OK);
		return -1;
	}

	// WinSockの初期化
	WSADATA wsaData;
	if(WSAStartup(MAKEWORD(2,0), &wsaData) != 0){
		MessageBox(NULL, _T("WSAStartup failed."), _T("Error"), MB_OK);
		return -1;
	}

	// メインダイアログ作成
	MainDialogBox dlg(hInst, devices);
	HWND hWnd = dlg.Create();
	if(!hWnd){
		return -1;
	}
	
	// メッセージループ
	MSG msg;
	while(GetMessage(&msg, NULL, 0, 0) != 0){
		if(!IsDialogMessage(hWnd, &msg)){
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
	}

	// WinSockの解放
	WSACleanup();

	return (int)msg.wParam;
}
