/**
 * @file
 * @brief Strawberry Linux USB���x�E���x�v���W���[���E�L�b�g(���[�J�[�i�ԁFUSBRH)�p�̃��j�^�����O�c�[���ł��B
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
 * USBMeter.dll�ւ̃A�N�Z�X��񋟂���֐����܂Ƃ߂����O��Ԃł��B
 * ��ԍŏ��� usbmeter::LoadUSBMeterDLL ���Ăяo���Ă��������B
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
 * BSTR��������܂��BSysFreeString(p)�����s���܂��Bshared_ptr�p�̍폜�q�Ƃ��Ďg���܂��B
 */
void ReleaseBSTR(BSTR p) { ::SysFreeString(p); }


/**
 * �ڑ����Ă���usbrh�f�o�C�X�̃f�o�C�X����񋓂��܂��B
 */
void EnumerateDevices(std::vector<boost::shared_ptr<OLECHAR> > *devices)
{
	LONG index = 0;
	for(;;){
		BSTR dev = usbmeter::FindUSB(&index); //Unicode�r���h�ł�ANSI�ŕԂ��Ă���̂Œ��ӁB
		// CSimpleArray<CAdapt<CComBSTR> >�Ȃ񂩂��g���ׂ��Ȃ̂�������Ȃ������ATL�ɂ͏ڂ����Ȃ����R�s�[�Ƃ����C�ɂȂ�̂ŁA�Ƃ肠����std::vector��shared_ptr���g�����Ƃɂ���B
		// std::vector<BSTR>�ŉ�������͍s��Ȃ����Ă̂����肩���B�ǂ����f�o�C�X��񋓂���͈̂�񂾂������B
		boost::shared_ptr<OLECHAR> devptr(dev, ReleaseBSTR);
		if(::SysStringLen(dev) == 0){
			break;
		}
		devices->push_back(devptr);
	}
	return;
}









// --------------------------------------------------------------------------
// ���C���_�C�A���O�{�b�N�X
// --------------------------------------------------------------------------

/**
 * �A�v���P�[�V�����̒��S�ƂȂ�_�C�A���O�{�b�N�X�E�B���h�E�ł��B
 *
 * �_�C�A���O�ɂ͌��݂̌v���f�[�^��ݒ��\�����܂��B
 * �^�C�}�[�C�x���g�Ōv���⃍�O�o�͂��s���܂��B
 * �N�����̓^�X�N�A�C�R����o�^���Ĕ�\����ԂɂȂ�܂��B
 */
class MainDialogBox
{
	static const UINT WM_NOTIFY_ICON = (WM_USER + 100);///< �V�X�e���g���C�A�C�R������̒ʒm�p�B
	static const UINT TIMER_ID = 1; ///< ���j�^�X�V�A�o�ߎ��Ԋm�F�Ɏg�p����^�C�}�[��ID�B
	static const UINT TIMER_PERIOD = 1000; ///< ���j�^�X�V�p�x�A�o�ߎ��Ԋm�F�p�x�B

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
	 * �_�C�A���O���b�Z�[�W�����b�Z�[�W���e�֐��ɐU�蕪���邾���ł��B
	 * @return TRUE(1)�̂Ƃ����b�Z�[�W������
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
	 * WM_INITDIALOG���b�Z�[�W�̏���
	 */
	BOOL OnInitDialog(HWND hWnd, HWND hWndFocus, LPARAM lParam)
	{
		// �ݒ�̓ǂݍ���
		LoadSettingsFromRegistory();
		CopySettingsVariablesToDlgItems();

		// �|�b�v�A�b�v���j���[�쐬
		hPopupMenu_ = ::LoadMenu(hInst_, _T("POPUPMENU"));

		// �A�C�R���ǂݍ���
		hIconApp_ = ::LoadIcon(hInst_, _T("IDI_APP"));

		// �A�C�R���ݒ�
		::ZeroMemory(&NotifyIconData_, sizeof(NotifyIconData_));
		NotifyIconData_.cbSize = sizeof(NotifyIconData_);
		NotifyIconData_.hWnd   = hWnd;
		NotifyIconData_.uID    = NULL;
		NotifyIconData_.uFlags = NIF_MESSAGE | NIF_ICON | NIF_TIP;
		NotifyIconData_.uCallbackMessage = WM_NOTIFY_ICON;
		NotifyIconData_.hIcon  = hIconApp_;
		::lstrcpy(NotifyIconData_.szTip, _T("USBRH Thermo-Hygrometer"));

		::Shell_NotifyIcon(NIM_ADD, &NotifyIconData_);

		// �^�C�}�[�ݒ�
		::GetLocalTime(&prevTime_);
		::SetTimer(hWnd, TIMER_ID, TIMER_PERIOD, NULL);

		return TRUE;
	}

	/**
	 * WM_DESTROY���b�Z�[�W����
	 */
	int OnDestroy(HWND hWnd)
	{
		PostQuitMessage(0);

		// �|�b�v�A�b�v���j���[���폜
		DestroyMenu(hPopupMenu_);

		// �C���W�Q�[�^�̈悩��폜
		Shell_NotifyIcon(NIM_DELETE, &NotifyIconData_);

		// �^�C�}�[���폜
		KillTimer(hWnd, TIMER_ID);

		return TRUE;
	}

	/**
	 * �_�C�A���O�{�b�N�X��\����Ԃɂ��܂��B
	 */
	void Show(void)
	{
		if(!visible_){
			visible_ = true;
			ShowWindow(hWnd_, SW_SHOW);
		}
	}

	/**
	 * �_�C�A���O�{�b�N�X���\����Ԃɂ��܂��B
	 */
	void Hide(void)
	{
		if(visible_){
			visible_ = false;
			ShowWindow(hWnd_, SW_HIDE);
		}
	}

	/**
	 * WM_CLOSE���b�Z�[�W����
	 */
	int OnClose(HWND hWnd)
	{
		Hide();
		return TRUE;
	}

	/**
	 * WM_COMMAND���b�Z�[�W����
	 * @param hWnd �E�B���h�E�n���h��	
	 * @param id ����ID�A�R���g���[��ID�A�A�N�Z�����[�^ID
	 * @param hWndCtl �R���g���[���̃n���h��
	 * @param codeNotify �ʒm�R�[�h
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
				_T("�^�C�g��"),
				MB_OK);
			break;
		case IDM_EXIT:
			DestroyWindow(hWnd);
			break;
		// ���C���_�C�A���O����̃R�}���h
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
	 * �V�X�e���g���C���b�Z�[�W
	 * WM_NOTIFY_ICON���b�Z�[�W����
	 * @param hWnd �E�B���h�E�n���h��
	 * @param wParam
	 * @param lParam ���b�Z�[�W���
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


	// �ݒ�̕ۑ��Ɠǂݍ���
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
	 * �ݒ�������o�ϐ�����_�C�A���O�փR�s�[���܂��B
	 * ���W�X�g������ǂݍ��񂾌��L�����Z���̎��ȂǂɎg�p���܂��B
	 */
	void CopySettingsVariablesToDlgItems(void)
	{
		SetDlgItemText(hWnd_, IDC_EDIT_OUTPUT_FILE, outputFilename_.c_str());
		SetDlgItemText(hWnd_, IDC_EDIT_OUTPUT_URL, outputHTTPURL_.c_str());
	}

	/**
	 * �ݒ���_�C�A���O���烁���o�ϐ��փR�s�[���܂��B
	 * OK�{�^�����������Ƃ��Ɏg�p���܂��B
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


	// �^�C�}�[�ɂ��v���Əo��


	void OnTimer(HWND hWnd, UINT TimerId)
	{
		SYSTEMTIME currTime;
		::GetLocalTime(&currTime);
		SYSTEMTIME prevTime = prevTime_;
		prevTime_ = currTime;

		if(currTime.wMinute != prevTime.wMinute){ ///@todo ���݂͈ꕪ�Ԋu�Œ�ɂȂ��Ă���̂ŁA�ςɂ������B
			Log(currTime, &MainDialogBox::OutputToFile);
			Log(currTime, &MainDialogBox::OutputToHTTP);
		}

		if(visible_){ // �_�C�A���O��\�����Ă���Ƃ�����
			Log(currTime, &MainDialogBox::OutputToDialog);
		}
	}

	void Log(const SYSTEMTIME &currTime, void (MainDialogBox::*func)(const SYSTEMTIME &, double, double))
	{
		if(devices_.empty()){
			return;
		}

		///@todo �����̃f�o�C�X�ɑΉ�����Busbrh������������Ă��Ȃ��̂őΉ�����C�ɂȂ�Ȃ��B�ǂ������o�͂ɂ��ׂ������悭������Ȃ��B

		double temp = 0;
		double humid = 0;
		if(usbmeter::GetTempHumidTrue(devices_[0].get(), &temp, &humid) != 0){
			return;
		}

		(this->*func)(currTime, temp, humid);
	}

	/**
	 * �v���f�[�^���t�@�C���֒ǋL�o�͂��܂��B
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
	 * �v���f�[�^���_�C�A���O�{�b�N�X�֏o�͂��܂��B
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
	 * �v���f�[�^��HTTP��GET���\�b�h�̃p�����[�^�Ƃ��ďo�͂��܂��B
	 * http://example.jp/record.cgi?temp=25.543&humid=45.123 �̂悤�Ȍ`���ŃA�N�Z�X���܂��B
	 */
	void OutputToHTTP(const SYSTEMTIME &currTime, double temp, double humid)
	{
		if(outputHTTPURL_.empty()){
			return;
		}

		///@todo �|�[�g�w��ł���悤�ɂ���B
		///@todo �擪�� http:// �Ɠ���Ă����v�Ȃ悤�ɂ���B
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

			// �ڑ�
			if(Local::connectToHost(sock, destHostName.c_str())){

				// ���N�G�X�g�쐬
				char requestParams[1024];
				std::sprintf(requestParams, "?temp=%lf&humid=%lf", temp, humid);
				std::string request = std::string("GET ") + destURL + requestParams + " HTTP/1.0\r\n\r\n";

				// HTTP���N�G�X�g���M
				bool error = false;
				int n = send(sock, request.c_str(), request.size(), 0);
				if(n < 0) {
					error = true;
				}

				// HTTP���b�Z�[�W��M
				std::string result;
				while(n > 0){
					char buf[128];
					n = recv(sock, buf, sizeof(buf), 0);
					if(n < 0) {
						error = true;
						break;
					}

					// ��M����
					result.append(buf, n);
				}
			}

			closesocket(sock);
		}
	}

};// class MainDialogBox


int WINAPI WinMain(HINSTANCE hInst, HINSTANCE hPrevInst, LPSTR lpszCmdParam, int nCmdShow)
{
	// ���x�v�E���x�v�f�o�C�X�̏�����
	if(!usbmeter::LoadUSBMeterDLL()){
		return -1;
	}
	std::vector<boost::shared_ptr<OLECHAR> > devices;
	EnumerateDevices(&devices);
	if(devices.empty()){
		MessageBox(NULL, _T("No USB meter device found."), _T("Error"), MB_OK);
		return -1;
	}

	// WinSock�̏�����
	WSADATA wsaData;
	if(WSAStartup(MAKEWORD(2,0), &wsaData) != 0){
		MessageBox(NULL, _T("WSAStartup failed."), _T("Error"), MB_OK);
		return -1;
	}

	// ���C���_�C�A���O�쐬
	MainDialogBox dlg(hInst, devices);
	HWND hWnd = dlg.Create();
	if(!hWnd){
		return -1;
	}
	
	// ���b�Z�[�W���[�v
	MSG msg;
	while(GetMessage(&msg, NULL, 0, 0) != 0){
		if(!IsDialogMessage(hWnd, &msg)){
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
	}

	// WinSock�̉��
	WSACleanup();

	return (int)msg.wParam;
}
