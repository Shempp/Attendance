#include "Attendance.h"

HWND hWnd;
CRITICAL_SECTION cs;
ofstream logFile;
char readBuffer[256];

VOID Date(char* str , int size , const char* format)
{
	time_t seconds = time(NULL);
	tm* timeinfo = localtime(&seconds);
	strftime(str, size, format, timeinfo);
}

VOID ReadConfigFile(ComConfData *pointer)
{
	ifstream confFile;
	string message;
	const int bufSize = 30;
	char confBuf[bufSize] = {0};
	char date[20] = {0};

	confFile.open("conf.ini");
	if (!confFile.is_open())
	{
		message = "Error: Файл отсутствует в папке с программой или не может быть открыт: conf.ini "
				"\nНажмите OK для завершения программы.";
		MessageBox(NULL, message.c_str(), "Error", MB_OK | MB_ICONERROR);
		Date(date, sizeof(date));
		message = ": >Ошибка при чтении файла \"conf.ini\". Данный файл должен находиться в одной папке с программой. "
				"\nТакже необходимо удостовериться, что файл не поврежден.";
		logFile << date << message << endl;
		message = ": >Программа завершилась с ошибкой.";
		logFile << date << message;
		logFile.close();
		exit(0);
	}
	else
	{
		confFile.close();
		Date(date, sizeof(date));
		message = ": >Файл \"conf.ini\" найден и готов к работе.";
		logFile << date << message << endl;
	}

	GetPrivateProfileString("general", "room", "Error", confBuf, bufSize, ".\\conf.ini");
	if ((GetLastError()) || (confBuf[0] == 0) || (strlen(confBuf) > 3))
	{
		message = "Error: Ошибка при чтении файла: conf.ini \nСм. файл logfile.txt для подробной информации.";
		MessageBox(NULL, message.c_str(), "Error", MB_OK | MB_ICONERROR);
		Date(date, sizeof(date));
		message = ": >Ошибка при чтении файла \"conf.ini\". Переменная \"room\" секции \"[general]\" либо не определена, "
				"либо количество символов в переменной превышает допустимое значение (максимум 3 символа). "
				"\nПример использования данной переменной: room=206";
		logFile << date << message << endl;
		Date(date, sizeof(date));
		message = ": >Программа завершилась с ошибкой.";
		logFile << date << message;
		logFile.close();
		exit(0);
	}
	else 
	{ 
		strncpy(pointer->room, confBuf, sizeof(confBuf));
		memset(confBuf, 0, sizeof(confBuf));
		Date(date, sizeof(date));
		message = ": >Переменная \"room\" секции \"[general]\" файла \"conf.ini\" определена.";
		logFile << date << message << endl;
	}

	GetPrivateProfileString("general", "building", "Error", confBuf, bufSize, ".\\conf.ini");
	if ((GetLastError()) || (confBuf[0] == 0) || (strlen(confBuf) > 1))
	{
		message = "Error: Ошибка при чтении файла: conf.ini \nСм. файл logfile.txt для подробной информации.";
		MessageBox(NULL, message.c_str(), "Error", MB_OK | MB_ICONERROR);
		Date(date, sizeof(date));
		message = ": >Ошибка при чтении файла \"conf.ini\". Переменная \"building\" секции \"[general]\" либо не определена, "
					"либо количество символов в переменной превышает допустимое значение (максимум 1 символ). "
					"\nПример использования данной переменной: building=1";
		logFile << date << message << endl;
		Date(date, sizeof(date));
		message = ": >Программа завершилась с ошибкой.";
		logFile << date << message;
		logFile.close();
		exit(0);
	}
	else 
	{
		strncpy(pointer->building, confBuf, sizeof(confBuf));
		memset(confBuf, 0, sizeof(confBuf));
		Date(date, sizeof(date));
		message = ": >Переменная \"building\" секции \"[general]\" файла \"conf.ini\" определена.";
		logFile << date << message << endl;
	}

	GetPrivateProfileString("general", "teacher", "Error", confBuf, bufSize, ".\\conf.ini");
	if ((GetLastError()) || (confBuf[0] == 0) || (strlen(confBuf) > 10))
	{
		message = "Error: Ошибка при чтении файла: conf.ini \nСм. файл logfile.txt для подробной информации.";
		MessageBox(NULL, message.c_str(), "Error", MB_OK | MB_ICONERROR);
		Date(date, sizeof(date));
		message = ": >Ошибка при чтении файла \"conf.ini\". Переменная \"teacher\" секции \"[general]\" либо не определена, "
				"либо количество символов в переменной превышает допустимое значение (максимум 10 символов). "
				"\nПример использования данной переменной: teacher=Петров";
		logFile << date << message << endl;
		Date(date, sizeof(date));
		message = ": >Программа завершилась с ошибкой.";
		logFile << date << message;
		logFile.close();
		exit(0);
	}
	else
	{
		MultiByteToWideChar(CP_ACP, 0, confBuf, -1, pointer->teacher, sizeof(pointer->teacher));
		memset(confBuf, 0, sizeof(confBuf));
		Date(date, sizeof(date));
		message = ": >Переменная \"teacher\" секции \"[general]\" файла \"conf.ini\" определена.";
		logFile << date << message << endl;
	}

	GetPrivateProfileString("general", "discipline", "Error", confBuf, bufSize, ".\\conf.ini");
	if ((GetLastError()) || (confBuf[0] == 0) || (strlen(confBuf) > 30))
	{
		message = "Error: Ошибка при чтении файла: conf.ini \nСм. файл logfile.txt для подробной информации.";
		MessageBox(NULL, message.c_str(), "Error", MB_OK | MB_ICONERROR);
		Date(date, sizeof(date));
		message = ": >Ошибка при чтении файла \"conf.ini\". Переменная \"discipline\" секции \"[general]\" либо не определена, "
				"либо количество символов в переменной превышает допустимое значение (максимум 30 символов). "
				"\nПример использования данной переменной: discipline=Компьютерные вирусы";
		logFile << date << message << endl;
		Date(date, sizeof(date));
		message = ": >Программа завершилась с ошибкой.";
		logFile << date << message;
		logFile.close();
		exit(0);
	}
	else
	{
		MultiByteToWideChar(CP_ACP, 0, confBuf, -1, pointer->discipline, sizeof(pointer->discipline));
		memset(confBuf, 0, sizeof(confBuf));
		Date(date, sizeof(date));
		message = ": >Переменная \"discipline\" секции \"[general]\" файла \"conf.ini\" определена.";
		logFile << date << message << endl;
	}

	GetPrivateProfileString("general", "subject", "Error", confBuf, bufSize, ".\\conf.ini");
	if ((GetLastError()) || (confBuf[0] == 0) || (strlen(confBuf) > 20))
	{
		message = "Error: Ошибка при чтении файла: conf.ini \nСм. файл logfile.txt для подробной информации.";
		MessageBox(NULL, message.c_str(), "Error", MB_OK | MB_ICONERROR);
		Date(date, sizeof(date));
		message = ": >Ошибка при чтении файла \"conf.ini\". Переменная \"subject\" секции \"[general]\" либо не определена, "
					"либо количество символов в переменной превышает допустимое значение (максимум 20 символов). "
					"\nПример использования данной переменной: subject=Лекция";
		logFile << date << message << endl;
		Date(date, sizeof(date));
		message = ": >Программа завершилась с ошибкой.";
		logFile << date << message;
		logFile.close();
		exit(0);
	}
	else
	{
		MultiByteToWideChar(CP_ACP, 0, confBuf, -1, pointer->subject, sizeof(pointer->subject));
		memset(confBuf, 0, sizeof(confBuf));
		Date(date, sizeof(date));
		message = ": >Переменная \"subject\" секции \"[general]\" файла \"conf.ini\" определена.";
		logFile << date << message << endl;
	}
	
	GetPrivateProfileString("general", "group", "Error", confBuf, bufSize, ".\\conf.ini");
	if ((GetLastError()) || (confBuf[0] == 0) || (strlen(confBuf) > 10))
	{
		message = "Error: Ошибка при чтении файла: conf.ini \nСм. файл logfile.txt для подробной информации.";
		MessageBox(NULL, message.c_str(), "Error", MB_OK | MB_ICONERROR);
		Date(date, sizeof(date));
		message = ": >Ошибка при чтении файла \"conf.ini\". Переменная \"group\" секции \"[general]\" либо не определена, "
				"либо количество символов в переменной превышает допустимое значение (максимум 10 символов). "
				"\nПример использования данной переменной: subject=Лекция";
		logFile << date << message << endl;
		Date(date, sizeof(date));
		message = ": >Программа завершилась с ошибкой.";
		logFile << date << message;
		logFile.close();
		exit(0);
	}
	else
	{
		MultiByteToWideChar(CP_ACP, 0, confBuf, -1, pointer->group, sizeof(pointer->group));
		memset(confBuf, 0, sizeof(confBuf));
		Date(date, sizeof(date));
		message = ": >Переменная \"group\" секции \"[general]\" файла \"conf.ini\" определена.";
		logFile << date << message << endl;
	}
	
	GetPrivateProfileString("settings", "comport", "Error", confBuf, bufSize, ".\\conf.ini");
	if ((GetLastError()) || (confBuf[0] == 0) || (strlen(confBuf) > 5))
	{
		message = "Error: Ошибка при чтении файла: conf.ini \nСм. файл logfile.txt для подробной информации.";
		MessageBox(NULL, message.c_str(), "Error", MB_OK | MB_ICONERROR);
		Date(date, sizeof(date));
		message = ": >Ошибка при чтении файла \"conf.ini\". Переменная \"comport\" секции \"[settings]\" либо не определена, "
			"либо количество символов в переменной превышает допустимое значение (максимум 5 символов). "
			"\nПример использования данной переменной: comport=COM1";
		logFile << date << message << endl;
		message = "Чтобы узнать COM-порт, к которому подключено считывающее устройство, "
			"зайдите в Пуск->Панель управления->Диспетчер устройств->Порты (COM и LPT). "
			"В данной секции и будет находиться информация о текущем COM-порте.";
		logFile << message << endl;
		Date(date, sizeof(date));
		message = ": >Программа завершилась с ошибкой.";
		logFile << date << message;
		logFile.close();
		exit(0);
	}
	else
	{
		pointer->comPort = (string)confBuf;
		memset(confBuf, 0, sizeof(confBuf));
		Date(date, sizeof(date));
		message = ": >Переменная \"comport\" секции \"[settings]\" файла \"conf.ini\" определена.";
		logFile << date << message << endl;
	}

	int conf = GetPrivateProfileInt("settings", "baudrate", -1, ".\\conf.ini");
	if ((conf == -1))
	{
		message = "Error: Ошибка при чтении файла: conf.ini \nСм. файл logfile.txt для подробной информации.";
		MessageBox(NULL, message.c_str(), "Error", MB_OK | MB_ICONERROR);
		Date(date, sizeof(date));
		message = ": >Ошибка при чтении файла \"conf.ini\". Переменная \"baudrate\" секции \"[settings]\" не определена либо отсутствует."
			"\nПример использования данной переменной: baudrate=9600";
		logFile << date << message << endl;
		Date(date, sizeof(date));
		message = ": >Программа завершилась с ошибкой.";
		logFile << date << message;
		logFile.close();
		exit(0);
	}
	else 
	{
		pointer->baudRate = conf;
		Date(date, sizeof(date));
		message = ": >Переменная \"baudrate\" секции \"[settings]\" файла \"conf.ini\" определена.";
		logFile << date << message << endl;
	}

	conf = GetPrivateProfileInt("settings", "bytesize", -1, ".\\conf.ini");
	if (conf == -1)
	{
		message = "Error: Ошибка при чтении файла: conf.ini \nСм. файл logfile.txt для подробной информации.";
		MessageBox(NULL, message.c_str(), "Error", MB_OK | MB_ICONERROR);
		Date(date, sizeof(date));
		message = ": >Ошибка при чтении файла \"conf.ini\". Переменная \"bytesize\" секции \"[settings]\" либо не определена, "
			"либо отсутствует. \nПример использования данной переменной: bytesize=8";
		logFile << date << message << endl;
		Date(date, sizeof(date));
		message = ": >Программа завершилась с ошибкой.";
		logFile << date << message;
		logFile.close();
		exit(0);
	}
	else 
	{
		pointer->byteSize = conf;
		Date(date, sizeof(date));
		message = ": >Переменная \"bytesize\" секции \"[settings]\" файла \"conf.ini\" определена.";
		logFile << date << message << endl;
	}

	conf = GetPrivateProfileInt("settings", "parity", -1, ".\\conf.ini");
	if ((conf > 5) || (conf < 1))
	{
		message = "Error: Ошибка при чтении файла: conf.ini \nСм. файл logfile.txt для подробной информации.";
		MessageBox(NULL, message.c_str(), "Error", MB_OK | MB_ICONERROR);
		Date(date, sizeof(date));
		message = ": >Ошибка при чтении файла \"conf.ini\". Переменная \"parity\" секции \"[settings]\" либо не определена, "
			"либо отсутствует (переменная принимает целочисленные значения от 1 до 5). "
			"\nПример использования данной переменной: parity=3";
		logFile << date << message << endl;
		Date(date, sizeof(date));
		message = ": >Программа завершилась с ошибкой.";
		logFile << date << message;
		logFile.close();
		exit(0);
	}
	else 
	{
		pointer->parity = conf;
		Date(date, sizeof(date));
		message = ": >Переменная \"parity\" секции \"[settings]\" файла \"conf.ini\" определена.";
		logFile << date << message << endl;
	}

	conf = GetPrivateProfileInt("settings", "stopbits", -1, ".\\conf.ini");
	if ((conf > 3) || (conf < 1))
	{
		message = "Error: Ошибка при чтении файла: conf.ini \nСм. файл logfile.txt для подробной информации.";
		MessageBox(NULL, message.c_str(), "Error", MB_OK | MB_ICONERROR);
		Date(date, sizeof(date));
		message = ": >Ошибка при чтении файла \"conf.ini\". Переменная \"stopbits\" секции \"[settings]\" либо не определена, "
			"либо отсутствует (переменная принимает целочисленные значения от 1 до 3). "
			"\nПример использования данной переменной: stopbits=1";
		logFile << date << message << endl;
		Date(date, sizeof(date));
		message = ": >Программа завершилась с ошибкой.";
		logFile << date << message;
		logFile.close();
		exit(0);
	}
	else 
	{
		pointer->stopBits = conf;
		Date(date, sizeof(date));
		message = ": >Переменная \"stopbits\" секции \"[settings]\" файла \"conf.ini\" определена.";
		logFile << date << message << endl;
	}

	conf = GetPrivateProfileInt("optional settings", "flag", -1, ".\\conf.ini");
	if ((conf > 1) || (conf <= -1))
	{
		Date(date, sizeof(date));
		message = ": >ПРЕДУПРЕЖДЕНИЕ: Переменная \"flag\" секции \"[optional settings]\" либо не определена, либо отсутствует. "
			"По умолчанию второе устройство не будет подключено. Если вы хотите включить второе устройство: flag=1";
		logFile << date << message << endl;
	}
	else
	{
		pointer->secDevActive = conf;
		if (!(pointer->secDevActive == 0))
		{
			GetPrivateProfileString("optional settings", "comport2", "Error", confBuf, bufSize, ".\\conf.ini");
			if ((GetLastError()) || (confBuf[0] == 0) || (strlen(confBuf) > 5))
			{
				Date(date, sizeof(date));
				message = ": >ПРЕДУПРЕЖДЕНИЕ: Переменная \"comport2\" секции \"[optional settings]\" либо не определена, "
					"либо количество символов в переменной превышает допустимое значение (максимум 5 символов). "
					"По этой причине могу возникнуть ошибки при работе со вторым устройством.";
				logFile << date << message << endl;
				return;
			}
			else
			{
				pointer->comPort2 = (string)confBuf;
				memset(confBuf, 0, sizeof(confBuf));
				Date(date, sizeof(date));
				message = ": >Переменная \"comport2\" секции \"[optional settings]\" файла \"conf.ini\" определена.";
				logFile << date << message << endl;
			}

			conf = GetPrivateProfileInt("optional settings", "baudrate2", -1, ".\\conf.ini");
			if (conf == -1)
			{
				Date(date, sizeof(date));
				message = ": >ПРЕДУПРЕЖДЕНИЕ: Переменная \"baudrate2\" секции \"[optional settings]\" либо не определена, "
					"либо отсутствует. По этой причине могу возникнуть ошибки при работе со вторым устройством.";
				logFile << date << message << endl;
				return;
			}
			else 
			{
				pointer->baudRate2 = conf;
				Date(date, sizeof(date));
				message = ": >Переменная \"baudrate2\" секции \"[optional settings]\" файла \"conf.ini\" определена.";
				logFile << date << message << endl;
			}

			conf = GetPrivateProfileInt("optional settings", "bytesize2", -1, ".\\conf.ini");
			if (conf == -1)
			{
				Date(date, sizeof(date));
				message = ": >ПРЕДУПРЕЖДЕНИЕ: Переменная \"bytesize2\" секции \"[optional settings]\" либо не определена, "
					"либо отсутствует. По этой причине могу возникнуть ошибки при работе со вторым устройством.";
				logFile << date << message << endl;
				return;
			}
			else 
			{
				pointer->byteSize2 = conf;
				Date(date, sizeof(date));
				message = ": >Переменная \"bytesize2\" секции \"[optional settings]\" файла \"conf.ini\" определена.";
				logFile << date << message << endl;
			}

			conf = GetPrivateProfileInt("optional settings", "parity2", -1, ".\\conf.ini");
			if ((conf > 5) || (conf < 1))
			{
				Date(date, sizeof(date));
				message = ": >ПРЕДУПРЕЖДЕНИЕ: Переменная \"parity2\" секции \"[optional settings]\" либо не определена, "
					"либо отсутствует (переменная принимает целочисленные значения от 1 до 5). "
					"По этой причине могу возникнуть ошибки при работе со вторым устройством.";
				logFile << date << message << endl;
				return;
			}
			else 
			{
				pointer->parity2 = conf;
				Date(date, sizeof(date));
				message = ": >Переменная \"parity2\" секции \"[optional settings]\" файла \"conf.ini\" определена.";
				logFile << date << message << endl;
			}

			conf = GetPrivateProfileInt("optional settings", "stopbits2", -1, ".\\conf.ini");
			if (conf == -1 || (conf > 3) || (conf < 1))
			{
				Date(date, sizeof(date));
				message = ": >ПРЕДУПРЕЖДЕНИЕ: Переменная \"stopbits2\" секции \"[optional settings]\" либо не определена, "
					"либо отсутствует (переменная принимает целочисленные значения от 1 до 3). "
					"По этой причине могу возникнуть ошибки при работе со вторым устройством.";
				logFile << date << message << endl;
				return;
			}
			else 
			{
				pointer->stopBits2 = conf;
				Date(date, sizeof(date));
				message = ": >Переменная \"stopbits2\" секции \"[optional settings]\" файла \"conf.ini\" определена.";
				logFile << date << message << endl;
			}
		}
		else
		{
			Date(date, sizeof(date));
			message = ": >ПРЕДУПРЕЖДЕНИЕ: Второе устройство отключено. Если вы хотите включить второе устройство: flag=1";
			logFile << date << message << endl;
			return;
		}
	}
}

VOID Disconnect(HANDLE &COM, int numDevice) 
{
	char date[20];

	if(COM != INVALID_HANDLE_VALUE) 
	{
		COM = INVALID_HANDLE_VALUE;
		CloseHandle(COM);
 	}
	Date(date, sizeof(date));
	logFile << date << " Устройство" << numDevice << ": >Закрытие предыдущей сессии работы с последовательным портом." << endl;
}

VOID Connect(ComConfData *pointer, int numDevice) 
{
	DCB dcb; 
	COMMTIMEOUTS ct; 
	HANDLE COM;
	string comPort;
	string message;
	char date[20] = {0};
	int baudRate;
	int byteSize;
	int parity;
	int stopBits;

	if (numDevice == 1)
	{
		comPort = pointer->comPort;
		baudRate = pointer->baudRate;
		byteSize = pointer->byteSize;
		parity = pointer->parity;
		stopBits = pointer->stopBits;
	}
	if (numDevice == 2)
	{
		comPort = pointer->comPort2;
		baudRate = pointer->baudRate2;
		byteSize = pointer->byteSize2;
		parity = pointer->parity2;
		stopBits = pointer->stopBits2;
	}

	COM = CreateFile(comPort.c_str(), GENERIC_READ, 0, NULL, OPEN_EXISTING, FILE_FLAG_OVERLAPPED, NULL);
	
 	if(COM == INVALID_HANDLE_VALUE) 
	{
		CloseHandle(COM);
		if (numDevice == 1)
		{
			message = "Error: Введено неверное значение COM-порта в конфигурационном файле, "
				"либо считывающее устройство не подключено (устройство номер 1). \nСм. файл logfile.txt для подробной информации.";
			MessageBox(NULL, message.c_str(), "Error", MB_OK | MB_ICONERROR);
			Date(date, sizeof(date));
			message = ": >Ошибка: Введено неверное значение COM-порта в конфигурационном файле или считывающее устройство не подключено.";
			logFile << date << " Устройство" << numDevice << message << endl;
			message = "Чтобы узнать COM-порт, к которому подключено считывающее устройство, "
				"зайдите в Пуск->Панель управления->Диспетчер устройств->Порты (COM и LPT). "
				"В данной секции и будет находиться информация о текущем COM-порте.";
			logFile << message << endl;
			message = "Пример использования данной переменной: comport=COM1";
			logFile << message << endl;
			Date(date, sizeof(date));
			message = ": >Программа завершилась с ошибкой.";
			logFile << date << message;
			logFile.close();
			exit(0);
		}
		if (numDevice == 2)
		{
			message = "Error: Введено неверное значение COM-порта в конфигурационном файле, "
				"либо считывающее устройство не подключено (устройство номер 2). \nСм. файл logfile.txt для подробной информации.";
			MessageBox(NULL, message.c_str(), "Warning", MB_OK | MB_ICONWARNING);
			Date(date, sizeof(date));
			message = ": >Ошибка: Введено неверное значение COM-порта в конфигурационном файле или считывающее устройство не подключено.";
			logFile << date << " Устройство" << numDevice << message << endl;
			message = "Чтобы узнать COM-порт, к которому подключено считывающее устройство, "
				"зайдите в Пуск->Панель управления->Диспетчер устройств->Порты (COM и LPT). "
				"В данной секции и будет находиться информация о текущем COM-порте.";
			logFile << message << endl;
			message = "Пример использования данной переменной: comport2=COM1";
			logFile << message << endl;
			return;
		}
 	}
	Date(date, sizeof(date));
    logFile << date << " Устройство" << numDevice << ": >Устройство успешно подключено." << endl;

 	if (!GetCommState(COM, &dcb))
	{
		COM = INVALID_HANDLE_VALUE;
		CloseHandle(COM);
		if (numDevice == 1)
		{
			message = "Error: Не удалось извлечь данные о текущих настройках управляющих сигналов для "
				"коммуникационного устройства (устройство номер 1). \nСм. файл logfile.txt для подробной информации.";
			MessageBox(NULL, message.c_str(), "Error", MB_OK | MB_ICONERROR);
			Date(date, sizeof(date));
			message = ": >Ошибка: Не удалось извлечь данные о текущих настройках управляющих сигналов "
				"для указанного коммуникационного устройства. \nОбычно данная ошибка появляется "
				"при неполадках с самим считвающим устройством или при неполадках с операционной системой.";
			logFile << date << " Устройство" << numDevice << message << endl;
			message = "Попробуйте перезагрузить ПК, также необходимо отключить и снова подключить считывающее устройство. "
				"Если данные действия не помогли, необходимо поменять считывающее устройство.";
			logFile << message << endl;
			Date(date, sizeof(date));
			logFile << date << ": >Программа завершилась с ошибкой.";
			logFile.close();
			exit(0);
		}
		if (numDevice == 2)
		{
			message = "Error: Не удалось извлечь данные о текущих настройках управляющих сигналов для "
				"коммуникационного устройства (устройство номер 2). \nСм. файл logfile.txt для подробной информации.";
			MessageBox(NULL, message.c_str(), "Warning", MB_OK | MB_ICONWARNING);
			Date(date, sizeof(date));
			message = ": >Ошибка: Не удалось извлечь данные о текущих настройках управляющих сигналов "
				"для указанного коммуникационного устройства. \nОбычно данная ошибка появляется при неполадках "
				"с самим считвающим устройством или при неполадках с операционной системой.";
			logFile << date << " Устройство" << numDevice << message << endl;
			message = "Попробуйте перезагрузить ПК, также необходимо отключить и снова подключить "
				"считывающее устройство. Если данные действия не помогли, необходимо поменять считывающее устройство.";
			logFile << message << endl;
			return;
		}
	}
	Date(date, sizeof(date));
	message = ": >Извлечение данных о текущих настройках управляющих сигналов для коммуникационного устройства произошло успешно.";
	logFile << date << " Устройство" << numDevice << message << endl;
	
	dcb.DCBlength = sizeof(DCB);
 	dcb.BaudRate = DWORD(baudRate); 
 	dcb.ByteSize = byteSize; 

	switch(parity) 
	{
	case 1:dcb.Parity = EVENPARITY; 
		break;
	case 2:dcb.Parity = MARKPARITY; 
		break;
	case 3:dcb.Parity = NOPARITY; 
		break;
	case 4:dcb.Parity = ODDPARITY; 
		break;
	case 5:dcb.Parity = SPACEPARITY; 
		break;
	}

	switch(stopBits)
	{
	case 1:dcb.StopBits = ONESTOPBIT;
		break;
	case 2:dcb.StopBits = ONE5STOPBITS;
		break;
	case 3:dcb.StopBits = TWOSTOPBITS;
		break;
	}

 	if(!SetCommState(COM, &dcb)) 
	{
		COM = INVALID_HANDLE_VALUE;
		CloseHandle(COM);
		if (numDevice == 1)
		{
			message = "Error: Введены неправильные настройки управления коммуникационным устройством или настройки "
				"вовсе не заданы (устройство номер 1). \nСм. файл logfile.txt для подробной информации.";
			MessageBox(NULL, message.c_str(), "Error", MB_OK | MB_ICONERROR);
			Date(date, sizeof(date));
			message = ": >Ошибка: Введены неправильные значения структуры DCB "
				"(настройки управления коммуникационным устройством) или значения вовсе не заданы.";
			logFile << date << " Устройство" << numDevice << message << endl;
			message = "Данная ошибка появляется при неправильной настройке считывающего устройства. "
				"Настройка производится из файла conf.ini "
				"\nРекомендуемые параметры: baudrate=9600, bytesize=8, parity=3(NOPARITY), stopbits=1(ONESTOPBIT).";
			logFile << message << endl;
			Date(date, sizeof(date));
			logFile << date << ": >Программа завершилась с ошибкой.";
			logFile.close();
			exit(0);
		}
		if (numDevice == 2)
		{
			message = "Error: Введены неправильные настройки управления коммуникационным устройством или "
				"настройки вовсе не заданы (устройство номер 2). \nСм. файл logfile.txt для подробной информации.";
			MessageBox(NULL, message.c_str(), "Warning", MB_OK | MB_ICONWARNING);
			Date(date, sizeof(date));
			message = ": >Ошибка: Введены неправильные значения структуры DCB "
				"(настройки управления коммуникационным устройством) или значения вовсе не заданы.";
			logFile << date << " Устройство" << numDevice << message << endl;
			message = "Данная ошибка появляется при неправильной настройке считывающего устройства. "
				"Настройка производится из файла conf.ini "
				"\nРекомендуемые параметры: baudrate2=9600, bytesize2=8, parity2=3(NOPARITY), stopbits2=1(ONESTOPBIT).";
			logFile << message << endl;
			return;
		}
 	}
	Date(date, sizeof(date));
	message = ": >Установка заданных настроек управления коммуникационным устройством произошла успешно.";
	logFile << date << " Устройство" << numDevice << message << endl;

	if(!SetCommMask(COM, EV_RXCHAR)) 
	{                                   
		COM = INVALID_HANDLE_VALUE;
		CloseHandle(COM);
		if (numDevice == 1)
		{
			message = "Error: Настройка пакета событий, которые будут отслеживаться для данного "
				"коммуникационного устройства невозможна (устройство номер 1). \nСм. файл logfile.txt для подробной информации.";
			MessageBox(NULL, message.c_str(), "Error", MB_OK | MB_ICONERROR);
			Date(date, sizeof(date));
			message = ": >Ошибка: Настройка маски EV_RXCHAR для данного коммуникационного устройства невозможна.";
			logFile << date << " Устройство" << numDevice << message << endl;
			message = "Ошибка при выполнении функции SetCommMask(COMPort, EV_RXCHAR) для отслеживания определенных "
				"событий от коммуникационного устройства. Маска EV_RXCHAR: символ был принят и помещен в буфер ввода данных. "
				"\nПопробуйте изменить настройки оборудования и перезапустить его. Если не помогает - необходимо поменять "
				"считывающее устройство.";
			logFile << message << endl;
			Date(date, sizeof(date));
			logFile << date << ": >Программа завершилась с ошибкой.";
			logFile.close();
			exit(0);
		}
		if (numDevice == 2)
		{
			message = "Error: Настройка пакета событий, которые будут отслеживаться для данного "
				"коммуникационного устройства невозможна (устройство номер 2). \nСм. файл logfile.txt для подробной информации.";
			MessageBox(NULL, message.c_str(), "Warning", MB_OK | MB_ICONWARNING);
			Date(date, sizeof(date));
			message = ": >Ошибка: Настройка маски EV_RXCHAR для данного коммуникационного устройства невозможна.";
			logFile << date << " Устройство" << numDevice << message << endl;
			message = "Ошибка при выполнении функции SetCommMask(COMPort, EV_RXCHAR) для отслеживания "
				"определенных событий от коммуникационного устройства. Маска EV_RXCHAR: символ был принят "
				"и помещен в буфер ввода данных. \nПопробуйте изменить настройки оборудования и перезапустить его. "
				"Если не помогает - необходимо поменять считывающее устройство.";
			logFile << message << endl;
			return;
		}
 	}
	Date(date, sizeof(date));
	message = ": >Настройка пакета событий, которые будут отслеживаться для данного коммуникационного устройства, произошла успешно.";
	logFile << date << " Устройство" << numDevice << message << endl;
	
 	ct.ReadIntervalTimeout = MAXDWORD; 
 	ct.ReadTotalTimeoutMultiplier = 0; 
 	ct.ReadTotalTimeoutConstant = 0; 
	
 	if(!SetCommTimeouts(COM, &ct)) 
	{
		COM = INVALID_HANDLE_VALUE;
		CloseHandle(COM);
		if (numDevice == 1)
		{
			message = "Error: Введены неправильные параметры интервала простоя по времени для коммуникационного "
				"устройства или параметры вовсе не заданы (устройство номер 1). "
				"\nСм. файл logfile.txt для подробной информации.";
			MessageBox(NULL, message.c_str(), "Error", MB_OK | MB_ICONERROR);
			Date(date, sizeof(date));
			message = ": >Ошибка: Введены неправильные значения структуры COMMTIMEOUTS "
				"(параметры интервала простоя по времени для коммуникационного устройства) или значения не заданы.";
			logFile << date << " Устройство" << numDevice << message << endl;
			message = "Попробуйте изменить настройки оборудования и перезапустить его. Если не "
				"помогает - необходимо поменять считывающее устройство.";
			logFile << message << endl;
			Date(date, sizeof(date));
			logFile << date << ": >Программа завершилась с ошибкой.";
			logFile.close();
			exit(0);
		}
		if (numDevice == 2)
		{
			message = "Error: Введены неправильные параметры интервала простоя по времени для коммуникационного "
				"устройства или параметры вовсе не заданы (устройство номер 2). "
				"\nСм. файл logfile.txt для подробной информации.";
			MessageBox(NULL, message.c_str(), "Warning", MB_OK | MB_ICONWARNING);
			Date(date, sizeof(date));
			message = ": >Ошибка: Введены неправильные значения структуры COMMTIMEOUTS "
				"(параметры интервала простоя по времени для коммуникационного устройства) или значения не заданы.";
			logFile << date << " Устройство" << numDevice << message << endl;
			message = "Попробуйте изменить настройки оборудования и перезапустить его. "
				"Если не помогает - необходимо поменять считывающее устройство.";
			logFile << message << endl;
			return;
		}
 	}

	if (numDevice == 1)
	{
		pointer->COM = COM;
	}
	if (numDevice == 2)
	{
		pointer->COM2 = COM;
	}

	Date(date, sizeof(date));
	message = ": >Настройка параметров интервала простоя по времени для коммуникационного устройства произошла успешно.";
	logFile << date << " Устройство" << numDevice << message << endl;
}

VOID WriteExcel(ComConfData *pointer)
{
	BasicExcelCell* cell;
	BasicExcelWorksheet* sheet = pointer->xls.GetWorksheet(0);
	XLSFormatManager fmtMgr(pointer->xls);
	string readId;
	static int row = 0; 
	static bool xlsCreate = true;
	static char excelName[256] = {0}; 
	char date[20] = {0};

	row ++;
	cell = sheet->Cell(row, 0);
	cell->SetString(readBuffer);

	char* dateFormat = "%d.%m.%Y";
	cell = sheet->Cell(row, 2);
	Date(date, sizeof(date), dateFormat);
	cell->SetString(date);

	char* timeFormat = "%H:%M:%S";
	cell = sheet->Cell(row, 3);
	Date(date, sizeof(date), timeFormat);
	cell->SetString(date);

	readId = (string)readBuffer;
	readId.pop_back();
	readId.pop_back();

	if (xlsCreate == true)
	{
		ExcelFont fontBold;
		CellFormat fmtBold(fmtMgr);
		int col;

		sheet->SetColWidth(0, 4000);
		sheet->SetColWidth(1, 4000);
		sheet->SetColWidth(2, 3000);
		sheet->SetColWidth(3, 3000);
		sheet->SetColWidth(4, 3000);
		sheet->SetColWidth(5, 2000);
		sheet->SetColWidth(6, 5000);
		sheet->SetColWidth(7, 7000);
		sheet->SetColWidth(8, 7000);

		fontBold._weight = FW_BOLD;
		fmtBold.set_font(fontBold);

		for (col = 0; col < 10; ++col)
		{
			cell = sheet->Cell(0, col);
			cell->SetFormat(fmtBold);
		}

		cell = sheet->Cell(0, 0);
		cell->SetString("ID");
		cell = sheet->Cell(0, 1);
		cell->SetWString(L"ФИО");
		cell = sheet->Cell(0, 2);
		cell->SetWString(L"Дата");
		cell = sheet->Cell(0, 3);
		cell->SetWString(L"Время");
		cell = sheet->Cell(0, 4);
		cell->SetWString(L"Аудитория");
		cell = sheet->Cell(0, 5);
		cell->SetWString(L"Корпус");
		cell = sheet->Cell(0, 6);
		cell->SetWString(L"Преподаватель (-и)");
		cell = sheet->Cell(0, 7);
		cell->SetWString(L"Дисциплина");
		cell = sheet->Cell(0, 8);
		cell->SetWString(L"Тип занятия");
		cell = sheet->Cell(0, 9);
		cell->SetWString(L"Группа");

		cell = sheet->Cell(1, 4);
		cell->Set(atoi(pointer->room));
		cell = sheet->Cell(1, 5);
		cell->Set(atoi(pointer->building));
		cell = sheet->Cell(1, 6);
		cell->SetWString(pointer->teacher);
		cell = sheet->Cell(1, 7);
		cell->SetWString(pointer->discipline);
		cell = sheet->Cell(1, 8);
		cell->SetWString(pointer->subject);
		cell = sheet->Cell(1, 9);
		cell->SetWString(pointer->group);

		char* format = "Day%d.%m.%Y Time%H.%M.%S";
		Date(excelName, sizeof(excelName), format);
		strcat(excelName, " Room");
		strcat(excelName, pointer->room);
		strcat(excelName, " Building");
		strcat(excelName, pointer->building);
		strcat(excelName, ".xls");
	}
	pointer->xls.SaveAs(excelName);

	if (xlsCreate == true)
	{
		Date(date, sizeof(date));
		logFile << date << ": >Создание Excel-файла произошло успешно." << endl;
	}
	xlsCreate = false;
	
	if (pointer->numDevice == 1)
	{
		Date(date, sizeof(date));
		logFile << date << " Устройство1" << ": >Идентификатор " << readId << " был успешно записан в Excel-файл." << endl;
	}
	if (pointer->numDevice == 2)
	{
		char* format = "%d.%m.%Y %H:%M:%S";
		Date(date, sizeof(date));
		logFile << date << " Устройство2" << ": >Идентификатор " << readId << " был успешно записан в Excel-файл." << endl;
	}
	
	SendMessage(hWnd, WM_TEXT, 0, 0);

	memset(readBuffer, 0, sizeof(readBuffer));	   
}

DWORD WINAPI ReadCom(ComConfData *pointer)
{
	OVERLAPPED overlapped; 
	COMSTAT comstat; 
	DWORD numBytes, temp, mask, signal;
	HANDLE COM = NULL;
	string message;
	char date[20];
	
	if (pointer->numDevice == 1)
	{
		COM = pointer->COM;
	}
	if (pointer->numDevice == 2)
	{
		COM = pointer->COM2;
	}
	
	memset(&overlapped, 0, sizeof(overlapped));
	overlapped.hEvent = CreateEvent(NULL, TRUE, TRUE, NULL);
	
	if (overlapped.hEvent == NULL)
	{
		message = "Error: При создании сигнального объекта для асинхронных операций возникла ошибка. "
			"\nСм. файл logfile.txt для подробной информации.";
		MessageBox(NULL, message.c_str(), "Error", MB_OK | MB_ICONERROR);
		Date(date, sizeof(date));
		message = ": >При создании сигнального объекта для асинхронных операций возникла ошибка.";
		logFile << date << message << endl;
		message = "Попробуйте перезагрузить ПК, также необходимо отключить и снова подключить считывающее устройство. "
			"Если данные действия не помогли, необходимо поменять считывающее устройство.";
		logFile << message << endl;
		Date(date, sizeof(date));
		logFile << date << ": >Программа завершилась с ошибкой.";
		logFile.close();
		exit(0);
	}

	while(1)				   
	{
		if (!WaitCommEvent(COM, &mask, &overlapped)) 
		{
			if (!GetLastError() == ERROR_IO_PENDING)
			{
				message = "Error: Функция WaitCommEvent завершилась с ошибкой. "
					"\nСм. файл logfile.txt для подробной информации.";
				MessageBox(NULL, message.c_str(), "Error", MB_OK | MB_ICONERROR);
				Date(date, sizeof(date));
				logFile << date << ": >Функция WaitCommEvent завершилась с ошибкой." << endl;
				message = "Попробуйте перезагрузить ПК, также необходимо отключить и снова подключить считывающее устройство. "
						"Если данные действия не помогли, необходимо поменять считывающее устройство.";
				logFile << message << endl;
				Date(date, sizeof(date));
				logFile << date << ": >Программа завершилась с ошибкой.";
				logFile.close();
				exit(0);
			}
		}

		signal = WaitForSingleObject(overlapped.hEvent, INFINITE);	
		if(signal == WAIT_OBJECT_0)				        
		{
			if(!GetOverlappedResult(COM, &overlapped, &temp, true))
			{
				message = "Error: Произошла ошибка в операции перекрытия. \nСм. файл logfile.txt для подробной информации.";
				MessageBox(NULL, message.c_str(), "Error", MB_OK | MB_ICONERROR);
				Date(date, sizeof(date));
				message = ": >Произошла ошибка в операции перекрытия, либо устройство было отключено во время работы программы.";
				logFile << date << message << endl;
				message = "Данную ошибку вызывает отключение устройства во время работы программы. "
					"Попробуйте перезагрузить ПК, также необходимо отключить и снова подключить считывающее устройство. "
					"Если данные действия не помогли, необходимо поменять считывающее устройство.";
				logFile << message << endl;
				Date(date, sizeof(date));
				logFile << date << ": >Программа завершилась с ошибкой.";
				logFile.close();
				exit(0);
			}
			if((mask & EV_RXCHAR) != 0) 
			{
				EnterCriticalSection(&cs);
				Sleep(200);
				ClearCommError(COM, &temp, &comstat); 
				numBytes = comstat.cbInQue; 
				if(numBytes) 
				{
					ReadFile(COM, readBuffer, numBytes, &temp, &overlapped); 
					WriteExcel(pointer); 
				}
				LeaveCriticalSection(&cs);
			}
		}
	}
}

VOID InitNotifyIconData(NOTIFYICONDATA &nid)
{
	TCHAR szTIP[64] = TEXT("Attendance");
    memset(&nid, 0, sizeof(NOTIFYICONDATA));
    nid.cbSize = sizeof(NOTIFYICONDATA);
    nid.hWnd = hWnd;
    nid.uID = ID_TRAY_APP_ICON;
    nid.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP| NIF_INFO;
    nid.uCallbackMessage = WM_SYSICON; 
    nid.hIcon = (HICON)LoadIcon (GetModuleHandle(NULL), MAKEINTRESOURCE(IDI_ICON1)) ;
    strncpy_s(nid.szTip, szTIP, sizeof(szTIP));
}