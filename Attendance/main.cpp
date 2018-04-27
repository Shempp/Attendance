#include "Attendance.h"

using namespace ExcelFormat;

HMENU hMenu;

LRESULT CALLBACK WindowProcedure(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	static NOTIFYICONDATA nid;
	TCHAR szInfo[256] = "Программа Attendance готова к работе."
						"\nЗапись индентификаторов будет производиться в Excel-файл, в одной папке с программой."
						"\nПодробный отчет о работе программы и об ошибках, возникающих во время работы программы, "
						"будет записываться в файл logfile.txt";
	TCHAR szInfoTitle[64] = "Внимание!";
	string titleBalloon = "Attendance";
	char textBalloon[256] = "Запись о пользователе ";
	char date[20];

    switch (message)                
    {
	case WM_ACTIVATE:
		InitNotifyIconData(nid);
		Shell_NotifyIcon(NIM_ADD, &nid);
		break;

    case WM_CREATE:
        hMenu = CreatePopupMenu();	
		AppendMenu(hMenu, MF_STRING, ID_TRAY_ABOUT,  TEXT("About"));
        AppendMenu(hMenu, MF_STRING, ID_TRAY_EXIT,  TEXT("Exit"));
        break;

	case WM_BEGIN:
		strncpy_s(nid.szInfo, szInfo, sizeof(szInfo));
		strncpy_s(nid.szInfoTitle, szInfoTitle, sizeof(szInfoTitle));
		Shell_NotifyIcon(NIM_MODIFY, &nid);
		break;

	case WM_TEXT:
		memset(szInfo, 0, sizeof(szInfo));
		memset(szInfoTitle, 0, sizeof(szInfoTitle));
		strcpy(szInfoTitle, titleBalloon.c_str());
		strcat(textBalloon, readBuffer);
		strcat(textBalloon, "успешно создана.");
		memcpy(szInfo, textBalloon, sizeof(readBuffer));
		strncpy_s(nid.szInfo, szInfo, sizeof(szInfo));
		strncpy_s(nid.szInfoTitle, szInfoTitle, sizeof(szInfoTitle));
		Shell_NotifyIcon(NIM_MODIFY, &nid);
		break;
   
    case WM_SYSICON: 
    {   
		switch(wParam)
        {
        case ID_TRAY_APP_ICON: 
			SetForegroundWindow(hWnd);
            break;
        }
		if (lParam == WM_LBUTTONUP)
		{
			HWND hWndWarning;
			hWndWarning = FindWindow(NULL, "Attendance");
			if(hWndWarning)
				DestroyWindow(hWndWarning);

			MessageBox(hWndWarning, "Приложите карту к считывающему устройству.", "Attendance", MB_OK | MB_ICONWARNING);
		}
        if (lParam == WM_RBUTTONDOWN) 
        {
            POINT curPoint;
            GetCursorPos(&curPoint);

            UINT clicked = TrackPopupMenu(hMenu, TPM_RETURNCMD | TPM_NONOTIFY, curPoint.x, curPoint.y, 0, hwnd, NULL); 
				  
            if (clicked == ID_TRAY_EXIT)
			{
				Date(date, sizeof(date));
                Shell_NotifyIcon(NIM_DELETE, &nid);
				logFile << date << ": >Работа программы завершена." << endl;
				logFile.close();
                PostQuitMessage(0);
            }
			if (clicked == ID_TRAY_ABOUT)
            {
				HWND hWndAbout;
				hWndAbout = FindWindow(NULL, "About Program");
				if(hWndAbout)
					DestroyWindow(hWndAbout);

				MessageBox(hWnd, "Attendance v1.3\noltirkov@mail.ru", "About Program", MB_OK | MB_ICONINFORMATION);
            }
        }
    }
    break;

    }

    return DefWindowProc(hwnd, message, wParam, lParam);
}

INT WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpszArgument, int nCmdShow)
{
	MSG msg; 
	WNDCLASSEX wincl = {0};
	ComConfData ccd;
	ComConfData ccdThread;
	ComConfData ccdThread2; 
	HANDLE thread;
	HANDLE thread2;
	string message;
	char szClassName[] = "Attendance";
	char date[20] = {0};

	Date(date, sizeof(date));
	logFile.open("logfile.txt");
	logFile << date <<  ": >Начало работы программы." << endl;

	ReadConfigFile(&ccd);
	ccd.xls.New(1);

	Disconnect(ccd.COM, 1);
	Connect(&ccd, 1);

	if (ccd.secDevActive == 1)
	{
		Disconnect(ccd.COM2, 2);
		Connect(&ccd, 2);
	}

	InitializeCriticalSection(&cs);

	ccdThread = ccd;
	ccdThread2 = ccd;

	ccdThread.numDevice = 1;
	thread = CreateThread(NULL, 0, (LPTHREAD_START_ROUTINE)ReadCom, &ccdThread, 0, NULL);
	if (thread == NULL)
	{
		message = "Error: Ошибка при создании отдельного потока для работы с коммуникационным оборудованием "
			"(устройство номер 1). \nСм. файл logfile.txt для подробной информации.";
		MessageBox(hWnd, message.c_str(), "Error", MB_OK | MB_ICONERROR);
		Date(date, sizeof(date));
		message = ": >Ошибка при создании отдельного потока для работы с коммуникационным оборудованием.";
		logFile << date << " Устройство1" << message << endl;
		message = "Попробуйте перезагрузить ПК, также необходимо отключить и снова "
			"подключить считывающее устройство. Если данные действия не помогли, "
			"необходимо поменять считывающее устройство. "
			"\nВ крайнем случае необходимы изменения в исходном коде программы.";
		logFile << message << endl;
		Date(date, sizeof(date));
		logFile << date << ": >Программа завершилась с ошибкой.";
		logFile.close();
		exit(0);
	}
	else
	{
		Date(date, sizeof(date));
		message = ": >Cоздание отдельного потока для работы с коммуникационным оборудованием произошло успешно.";
		logFile << date << " Устройство1" << message << endl;
	}
	
	ccdThread2.numDevice = 2;
	if (ccd.secDevActive == 1)
	{
		thread2 = CreateThread(NULL, 0, (LPTHREAD_START_ROUTINE)ReadCom, &ccdThread2, 0, NULL);
		if (thread2 == NULL)
		{
			message = "Error: Ошибка при создании отдельного потока для работы с коммуникационным "
				"оборудованием (устройство номер 2). \nСм. файл logfile.txt для подробной информации.";
			MessageBox(hWnd, message.c_str(), "Error", MB_OK | MB_ICONERROR);
			Date(date, sizeof(date));
			message = ": >Ошибка при создании отдельного потока для работы с коммуникационным оборудованием.";
			logFile << date << " Устройство2" << message << endl;
			message = "Попробуйте перезагрузить ПК, также необходимо отключить и снова подключить "
				"считывающее устройство. Если данные действия не помогли, необходимо поменять считывающее устройство. "
				"\nВ крайнем случае необходимы изменения в исходном коде программы.";
			logFile << message << endl;
		}
		else
		{
			Date(date, sizeof(date));
			message = ": >Cоздание отдельного потока для работы с коммуникационным оборудованием произошло успешно.";
			logFile << date << " Устройство2" << message << endl;
		}
	}
 
    wincl.hInstance = hInstance;
    wincl.lpszClassName = szClassName;
    wincl.lpfnWndProc = WindowProcedure;    
    wincl.style = NULL; 
    wincl.cbSize = sizeof(WNDCLASSEX);
    wincl.hIcon = LoadIcon (GetModuleHandle(NULL), MAKEINTRESOURCE(IDI_ICON1));
    wincl.hCursor = LoadCursor (NULL, IDC_ARROW);
    wincl.lpszMenuName = NULL; 
    wincl.cbClsExtra = 0;   
    wincl.cbWndExtra = 0;
    wincl.hbrBackground = NULL;

    if (!RegisterClassEx(&wincl))
	{
		message = "Error: Ошибка при регистрации класса окна. \nСм. файл logfile.txt для подробной информации.";
		MessageBox(hWnd, message.c_str(), "Error", MB_OK | MB_ICONERROR);
		Date(date, sizeof(date));
		logFile << date << ": >Ошибка при регистрации класса окна." << endl;
		message = "Попробуйте перезагрузить ПК, также необходимо отключить и снова подключить считывающее устройство. "
			"Если данные действия не помогли, необходимо поменять считывающее устройство. "
			"\nВ крайнем случае необходимы изменения в исходном коде программы.";
		logFile << message << endl;
		Date(date, sizeof(date));
		logFile << date << ": >Программа завершилась с ошибкой.";
		logFile.close();
		exit(0);
	}
	else
	{
		Date(date, sizeof(date));
		logFile << date << ": >Регистрация окна программы произошло успешно." << endl;
	}
	
    hWnd = CreateWindowEx(0, szClassName, NULL, 0, 0, 0, 0, 0, HWND_MESSAGE, NULL, hInstance, NULL);
	if (!hWnd)
	{
		message = "Error: Ошибка при создании окна программы. \nСм. файл logfile.txt для подробной информации.";
		MessageBox(hWnd, message.c_str(), "Error", MB_OK | MB_ICONERROR);
		Date(date, sizeof(date));
		logFile << date << ": >Ошибка при создании окна программы." << endl;
		message = "Попробуйте перезагрузить ПК, также необходимо отключить и снова подключить считывающее устройство. "
			"Если данные действия не помогли, необходимо поменять считывающее устройство. "
			"\nВ крайнем случае необходимы изменения в исходном коде программы.";
		logFile << message << endl;
		Date(date, sizeof(date));
		logFile << date <<  ": >Программа завершилась с ошибкой.";
		logFile.close();
		exit(0);
	}
	else
	{
		Date(date, sizeof(date));
		logFile << date <<  ": >Создание окна программы произошло успешно." << endl;
	}

    ShowWindow(hWnd, nCmdShow);
	UpdateWindow(hWnd);

	SendMessage(hWnd, WM_BEGIN, 0, 0);

    while(GetMessage(&msg, NULL, 0, 0))
    {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    return msg.wParam;
}

