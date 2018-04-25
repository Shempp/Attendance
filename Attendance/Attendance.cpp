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
		message = "Error: ���� ����������� � ����� � ���������� ��� �� ����� ���� ������: conf.ini "
				"\n������� OK ��� ���������� ���������.";
		MessageBox(NULL, message.c_str(), "Error", MB_OK | MB_ICONERROR);
		Date(date, sizeof(date));
		message = ": >������ ��� ������ ����� \"conf.ini\". ������ ���� ������ ���������� � ����� ����� � ����������. "
				"\n����� ���������� ��������������, ��� ���� �� ���������.";
		logFile << date << message << endl;
		message = ": >��������� ����������� � �������.";
		logFile << date << message;
		logFile.close();
		exit(0);
	}
	else
	{
		confFile.close();
		Date(date, sizeof(date));
		message = ": >���� \"conf.ini\" ������ � ����� � ������.";
		logFile << date << message << endl;
	}

	GetPrivateProfileString("general", "room", "Error", confBuf, bufSize, ".\\conf.ini");
	if ((GetLastError()) || (confBuf[0] == 0) || (strlen(confBuf) > 3))
	{
		message = "Error: ������ ��� ������ �����: conf.ini \n��. ���� logfile.txt ��� ��������� ����������.";
		MessageBox(NULL, message.c_str(), "Error", MB_OK | MB_ICONERROR);
		Date(date, sizeof(date));
		message = ": >������ ��� ������ ����� \"conf.ini\". ���������� \"room\" ������ \"[general]\" ���� �� ����������, "
				"���� ���������� �������� � ���������� ��������� ���������� �������� (�������� 3 �������). "
				"\n������ ������������� ������ ����������: room=206";
		logFile << date << message << endl;
		Date(date, sizeof(date));
		message = ": >��������� ����������� � �������.";
		logFile << date << message;
		logFile.close();
		exit(0);
	}
	else 
	{ 
		strncpy(pointer->room, confBuf, sizeof(confBuf));
		memset(confBuf, 0, sizeof(confBuf));
		Date(date, sizeof(date));
		message = ": >���������� \"room\" ������ \"[general]\" ����� \"conf.ini\" ����������.";
		logFile << date << message << endl;
	}

	GetPrivateProfileString("general", "building", "Error", confBuf, bufSize, ".\\conf.ini");
	if ((GetLastError()) || (confBuf[0] == 0) || (strlen(confBuf) > 1))
	{
		message = "Error: ������ ��� ������ �����: conf.ini \n��. ���� logfile.txt ��� ��������� ����������.";
		MessageBox(NULL, message.c_str(), "Error", MB_OK | MB_ICONERROR);
		Date(date, sizeof(date));
		message = ": >������ ��� ������ ����� \"conf.ini\". ���������� \"building\" ������ \"[general]\" ���� �� ����������, "
					"���� ���������� �������� � ���������� ��������� ���������� �������� (�������� 1 ������). "
					"\n������ ������������� ������ ����������: building=1";
		logFile << date << message << endl;
		Date(date, sizeof(date));
		message = ": >��������� ����������� � �������.";
		logFile << date << message;
		logFile.close();
		exit(0);
	}
	else 
	{
		strncpy(pointer->building, confBuf, sizeof(confBuf));
		memset(confBuf, 0, sizeof(confBuf));
		Date(date, sizeof(date));
		message = ": >���������� \"building\" ������ \"[general]\" ����� \"conf.ini\" ����������.";
		logFile << date << message << endl;
	}

	GetPrivateProfileString("general", "teacher", "Error", confBuf, bufSize, ".\\conf.ini");
	if ((GetLastError()) || (confBuf[0] == 0) || (strlen(confBuf) > 10))
	{
		message = "Error: ������ ��� ������ �����: conf.ini \n��. ���� logfile.txt ��� ��������� ����������.";
		MessageBox(NULL, message.c_str(), "Error", MB_OK | MB_ICONERROR);
		Date(date, sizeof(date));
		message = ": >������ ��� ������ ����� \"conf.ini\". ���������� \"teacher\" ������ \"[general]\" ���� �� ����������, "
				"���� ���������� �������� � ���������� ��������� ���������� �������� (�������� 10 ��������). "
				"\n������ ������������� ������ ����������: teacher=������";
		logFile << date << message << endl;
		Date(date, sizeof(date));
		message = ": >��������� ����������� � �������.";
		logFile << date << message;
		logFile.close();
		exit(0);
	}
	else
	{
		MultiByteToWideChar(CP_ACP, 0, confBuf, -1, pointer->teacher, sizeof(pointer->teacher));
		memset(confBuf, 0, sizeof(confBuf));
		Date(date, sizeof(date));
		message = ": >���������� \"teacher\" ������ \"[general]\" ����� \"conf.ini\" ����������.";
		logFile << date << message << endl;
	}

	GetPrivateProfileString("general", "discipline", "Error", confBuf, bufSize, ".\\conf.ini");
	if ((GetLastError()) || (confBuf[0] == 0) || (strlen(confBuf) > 30))
	{
		message = "Error: ������ ��� ������ �����: conf.ini \n��. ���� logfile.txt ��� ��������� ����������.";
		MessageBox(NULL, message.c_str(), "Error", MB_OK | MB_ICONERROR);
		Date(date, sizeof(date));
		message = ": >������ ��� ������ ����� \"conf.ini\". ���������� \"discipline\" ������ \"[general]\" ���� �� ����������, "
				"���� ���������� �������� � ���������� ��������� ���������� �������� (�������� 30 ��������). "
				"\n������ ������������� ������ ����������: discipline=������������ ������";
		logFile << date << message << endl;
		Date(date, sizeof(date));
		message = ": >��������� ����������� � �������.";
		logFile << date << message;
		logFile.close();
		exit(0);
	}
	else
	{
		MultiByteToWideChar(CP_ACP, 0, confBuf, -1, pointer->discipline, sizeof(pointer->discipline));
		memset(confBuf, 0, sizeof(confBuf));
		Date(date, sizeof(date));
		message = ": >���������� \"discipline\" ������ \"[general]\" ����� \"conf.ini\" ����������.";
		logFile << date << message << endl;
	}

	GetPrivateProfileString("general", "subject", "Error", confBuf, bufSize, ".\\conf.ini");
	if ((GetLastError()) || (confBuf[0] == 0) || (strlen(confBuf) > 20))
	{
		message = "Error: ������ ��� ������ �����: conf.ini \n��. ���� logfile.txt ��� ��������� ����������.";
		MessageBox(NULL, message.c_str(), "Error", MB_OK | MB_ICONERROR);
		Date(date, sizeof(date));
		message = ": >������ ��� ������ ����� \"conf.ini\". ���������� \"subject\" ������ \"[general]\" ���� �� ����������, "
					"���� ���������� �������� � ���������� ��������� ���������� �������� (�������� 20 ��������). "
					"\n������ ������������� ������ ����������: subject=������";
		logFile << date << message << endl;
		Date(date, sizeof(date));
		message = ": >��������� ����������� � �������.";
		logFile << date << message;
		logFile.close();
		exit(0);
	}
	else
	{
		MultiByteToWideChar(CP_ACP, 0, confBuf, -1, pointer->subject, sizeof(pointer->subject));
		memset(confBuf, 0, sizeof(confBuf));
		Date(date, sizeof(date));
		message = ": >���������� \"subject\" ������ \"[general]\" ����� \"conf.ini\" ����������.";
		logFile << date << message << endl;
	}
	
	GetPrivateProfileString("general", "group", "Error", confBuf, bufSize, ".\\conf.ini");
	if ((GetLastError()) || (confBuf[0] == 0) || (strlen(confBuf) > 10))
	{
		message = "Error: ������ ��� ������ �����: conf.ini \n��. ���� logfile.txt ��� ��������� ����������.";
		MessageBox(NULL, message.c_str(), "Error", MB_OK | MB_ICONERROR);
		Date(date, sizeof(date));
		message = ": >������ ��� ������ ����� \"conf.ini\". ���������� \"group\" ������ \"[general]\" ���� �� ����������, "
				"���� ���������� �������� � ���������� ��������� ���������� �������� (�������� 10 ��������). "
				"\n������ ������������� ������ ����������: subject=������";
		logFile << date << message << endl;
		Date(date, sizeof(date));
		message = ": >��������� ����������� � �������.";
		logFile << date << message;
		logFile.close();
		exit(0);
	}
	else
	{
		MultiByteToWideChar(CP_ACP, 0, confBuf, -1, pointer->group, sizeof(pointer->group));
		memset(confBuf, 0, sizeof(confBuf));
		Date(date, sizeof(date));
		message = ": >���������� \"group\" ������ \"[general]\" ����� \"conf.ini\" ����������.";
		logFile << date << message << endl;
	}
	
	GetPrivateProfileString("settings", "comport", "Error", confBuf, bufSize, ".\\conf.ini");
	if ((GetLastError()) || (confBuf[0] == 0) || (strlen(confBuf) > 5))
	{
		message = "Error: ������ ��� ������ �����: conf.ini \n��. ���� logfile.txt ��� ��������� ����������.";
		MessageBox(NULL, message.c_str(), "Error", MB_OK | MB_ICONERROR);
		Date(date, sizeof(date));
		message = ": >������ ��� ������ ����� \"conf.ini\". ���������� \"comport\" ������ \"[settings]\" ���� �� ����������, "
			"���� ���������� �������� � ���������� ��������� ���������� �������� (�������� 5 ��������). "
			"\n������ ������������� ������ ����������: comport=COM1";
		logFile << date << message << endl;
		message = "����� ������ COM-����, � �������� ���������� ����������� ����������, "
			"������� � ����->������ ����������->��������� ���������->����� (COM � LPT). "
			"� ������ ������ � ����� ���������� ���������� � ������� COM-�����.";
		logFile << message << endl;
		Date(date, sizeof(date));
		message = ": >��������� ����������� � �������.";
		logFile << date << message;
		logFile.close();
		exit(0);
	}
	else
	{
		pointer->comPort = (string)confBuf;
		memset(confBuf, 0, sizeof(confBuf));
		Date(date, sizeof(date));
		message = ": >���������� \"comport\" ������ \"[settings]\" ����� \"conf.ini\" ����������.";
		logFile << date << message << endl;
	}

	int conf = GetPrivateProfileInt("settings", "baudrate", -1, ".\\conf.ini");
	if ((conf == -1))
	{
		message = "Error: ������ ��� ������ �����: conf.ini \n��. ���� logfile.txt ��� ��������� ����������.";
		MessageBox(NULL, message.c_str(), "Error", MB_OK | MB_ICONERROR);
		Date(date, sizeof(date));
		message = ": >������ ��� ������ ����� \"conf.ini\". ���������� \"baudrate\" ������ \"[settings]\" �� ���������� ���� �����������."
			"\n������ ������������� ������ ����������: baudrate=9600";
		logFile << date << message << endl;
		Date(date, sizeof(date));
		message = ": >��������� ����������� � �������.";
		logFile << date << message;
		logFile.close();
		exit(0);
	}
	else 
	{
		pointer->baudRate = conf;
		Date(date, sizeof(date));
		message = ": >���������� \"baudrate\" ������ \"[settings]\" ����� \"conf.ini\" ����������.";
		logFile << date << message << endl;
	}

	conf = GetPrivateProfileInt("settings", "bytesize", -1, ".\\conf.ini");
	if (conf == -1)
	{
		message = "Error: ������ ��� ������ �����: conf.ini \n��. ���� logfile.txt ��� ��������� ����������.";
		MessageBox(NULL, message.c_str(), "Error", MB_OK | MB_ICONERROR);
		Date(date, sizeof(date));
		message = ": >������ ��� ������ ����� \"conf.ini\". ���������� \"bytesize\" ������ \"[settings]\" ���� �� ����������, "
			"���� �����������. \n������ ������������� ������ ����������: bytesize=8";
		logFile << date << message << endl;
		Date(date, sizeof(date));
		message = ": >��������� ����������� � �������.";
		logFile << date << message;
		logFile.close();
		exit(0);
	}
	else 
	{
		pointer->byteSize = conf;
		Date(date, sizeof(date));
		message = ": >���������� \"bytesize\" ������ \"[settings]\" ����� \"conf.ini\" ����������.";
		logFile << date << message << endl;
	}

	conf = GetPrivateProfileInt("settings", "parity", -1, ".\\conf.ini");
	if ((conf > 5) || (conf < 1))
	{
		message = "Error: ������ ��� ������ �����: conf.ini \n��. ���� logfile.txt ��� ��������� ����������.";
		MessageBox(NULL, message.c_str(), "Error", MB_OK | MB_ICONERROR);
		Date(date, sizeof(date));
		message = ": >������ ��� ������ ����� \"conf.ini\". ���������� \"parity\" ������ \"[settings]\" ���� �� ����������, "
			"���� ����������� (���������� ��������� ������������� �������� �� 1 �� 5). "
			"\n������ ������������� ������ ����������: parity=3";
		logFile << date << message << endl;
		Date(date, sizeof(date));
		message = ": >��������� ����������� � �������.";
		logFile << date << message;
		logFile.close();
		exit(0);
	}
	else 
	{
		pointer->parity = conf;
		Date(date, sizeof(date));
		message = ": >���������� \"parity\" ������ \"[settings]\" ����� \"conf.ini\" ����������.";
		logFile << date << message << endl;
	}

	conf = GetPrivateProfileInt("settings", "stopbits", -1, ".\\conf.ini");
	if ((conf > 3) || (conf < 1))
	{
		message = "Error: ������ ��� ������ �����: conf.ini \n��. ���� logfile.txt ��� ��������� ����������.";
		MessageBox(NULL, message.c_str(), "Error", MB_OK | MB_ICONERROR);
		Date(date, sizeof(date));
		message = ": >������ ��� ������ ����� \"conf.ini\". ���������� \"stopbits\" ������ \"[settings]\" ���� �� ����������, "
			"���� ����������� (���������� ��������� ������������� �������� �� 1 �� 3). "
			"\n������ ������������� ������ ����������: stopbits=1";
		logFile << date << message << endl;
		Date(date, sizeof(date));
		message = ": >��������� ����������� � �������.";
		logFile << date << message;
		logFile.close();
		exit(0);
	}
	else 
	{
		pointer->stopBits = conf;
		Date(date, sizeof(date));
		message = ": >���������� \"stopbits\" ������ \"[settings]\" ����� \"conf.ini\" ����������.";
		logFile << date << message << endl;
	}

	conf = GetPrivateProfileInt("optional settings", "flag", -1, ".\\conf.ini");
	if ((conf > 1) || (conf <= -1))
	{
		Date(date, sizeof(date));
		message = ": >��������������: ���������� \"flag\" ������ \"[optional settings]\" ���� �� ����������, ���� �����������. "
			"�� ��������� ������ ���������� �� ����� ����������. ���� �� ������ �������� ������ ����������: flag=1";
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
				message = ": >��������������: ���������� \"comport2\" ������ \"[optional settings]\" ���� �� ����������, "
					"���� ���������� �������� � ���������� ��������� ���������� �������� (�������� 5 ��������). "
					"�� ���� ������� ���� ���������� ������ ��� ������ �� ������ �����������.";
				logFile << date << message << endl;
				return;
			}
			else
			{
				pointer->comPort2 = (string)confBuf;
				memset(confBuf, 0, sizeof(confBuf));
				Date(date, sizeof(date));
				message = ": >���������� \"comport2\" ������ \"[optional settings]\" ����� \"conf.ini\" ����������.";
				logFile << date << message << endl;
			}

			conf = GetPrivateProfileInt("optional settings", "baudrate2", -1, ".\\conf.ini");
			if (conf == -1)
			{
				Date(date, sizeof(date));
				message = ": >��������������: ���������� \"baudrate2\" ������ \"[optional settings]\" ���� �� ����������, "
					"���� �����������. �� ���� ������� ���� ���������� ������ ��� ������ �� ������ �����������.";
				logFile << date << message << endl;
				return;
			}
			else 
			{
				pointer->baudRate2 = conf;
				Date(date, sizeof(date));
				message = ": >���������� \"baudrate2\" ������ \"[optional settings]\" ����� \"conf.ini\" ����������.";
				logFile << date << message << endl;
			}

			conf = GetPrivateProfileInt("optional settings", "bytesize2", -1, ".\\conf.ini");
			if (conf == -1)
			{
				Date(date, sizeof(date));
				message = ": >��������������: ���������� \"bytesize2\" ������ \"[optional settings]\" ���� �� ����������, "
					"���� �����������. �� ���� ������� ���� ���������� ������ ��� ������ �� ������ �����������.";
				logFile << date << message << endl;
				return;
			}
			else 
			{
				pointer->byteSize2 = conf;
				Date(date, sizeof(date));
				message = ": >���������� \"bytesize2\" ������ \"[optional settings]\" ����� \"conf.ini\" ����������.";
				logFile << date << message << endl;
			}

			conf = GetPrivateProfileInt("optional settings", "parity2", -1, ".\\conf.ini");
			if ((conf > 5) || (conf < 1))
			{
				Date(date, sizeof(date));
				message = ": >��������������: ���������� \"parity2\" ������ \"[optional settings]\" ���� �� ����������, "
					"���� ����������� (���������� ��������� ������������� �������� �� 1 �� 5). "
					"�� ���� ������� ���� ���������� ������ ��� ������ �� ������ �����������.";
				logFile << date << message << endl;
				return;
			}
			else 
			{
				pointer->parity2 = conf;
				Date(date, sizeof(date));
				message = ": >���������� \"parity2\" ������ \"[optional settings]\" ����� \"conf.ini\" ����������.";
				logFile << date << message << endl;
			}

			conf = GetPrivateProfileInt("optional settings", "stopbits2", -1, ".\\conf.ini");
			if (conf == -1 || (conf > 3) || (conf < 1))
			{
				Date(date, sizeof(date));
				message = ": >��������������: ���������� \"stopbits2\" ������ \"[optional settings]\" ���� �� ����������, "
					"���� ����������� (���������� ��������� ������������� �������� �� 1 �� 3). "
					"�� ���� ������� ���� ���������� ������ ��� ������ �� ������ �����������.";
				logFile << date << message << endl;
				return;
			}
			else 
			{
				pointer->stopBits2 = conf;
				Date(date, sizeof(date));
				message = ": >���������� \"stopbits2\" ������ \"[optional settings]\" ����� \"conf.ini\" ����������.";
				logFile << date << message << endl;
			}
		}
		else
		{
			Date(date, sizeof(date));
			message = ": >��������������: ������ ���������� ���������. ���� �� ������ �������� ������ ����������: flag=1";
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
	logFile << date << " ����������" << numDevice << ": >�������� ���������� ������ ������ � ���������������� ������." << endl;
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
			message = "Error: ������� �������� �������� COM-����� � ���������������� �����, "
				"���� ����������� ���������� �� ���������� (���������� ����� 1). \n��. ���� logfile.txt ��� ��������� ����������.";
			MessageBox(NULL, message.c_str(), "Error", MB_OK | MB_ICONERROR);
			Date(date, sizeof(date));
			message = ": >������: ������� �������� �������� COM-����� � ���������������� ����� ��� ����������� ���������� �� ����������.";
			logFile << date << " ����������" << numDevice << message << endl;
			message = "����� ������ COM-����, � �������� ���������� ����������� ����������, "
				"������� � ����->������ ����������->��������� ���������->����� (COM � LPT). "
				"� ������ ������ � ����� ���������� ���������� � ������� COM-�����.";
			logFile << message << endl;
			message = "������ ������������� ������ ����������: comport=COM1";
			logFile << message << endl;
			Date(date, sizeof(date));
			message = ": >��������� ����������� � �������.";
			logFile << date << message;
			logFile.close();
			exit(0);
		}
		if (numDevice == 2)
		{
			message = "Error: ������� �������� �������� COM-����� � ���������������� �����, "
				"���� ����������� ���������� �� ���������� (���������� ����� 2). \n��. ���� logfile.txt ��� ��������� ����������.";
			MessageBox(NULL, message.c_str(), "Warning", MB_OK | MB_ICONWARNING);
			Date(date, sizeof(date));
			message = ": >������: ������� �������� �������� COM-����� � ���������������� ����� ��� ����������� ���������� �� ����������.";
			logFile << date << " ����������" << numDevice << message << endl;
			message = "����� ������ COM-����, � �������� ���������� ����������� ����������, "
				"������� � ����->������ ����������->��������� ���������->����� (COM � LPT). "
				"� ������ ������ � ����� ���������� ���������� � ������� COM-�����.";
			logFile << message << endl;
			message = "������ ������������� ������ ����������: comport2=COM1";
			logFile << message << endl;
			return;
		}
 	}
	Date(date, sizeof(date));
    logFile << date << " ����������" << numDevice << ": >���������� ������� ����������." << endl;

 	if (!GetCommState(COM, &dcb))
	{
		COM = INVALID_HANDLE_VALUE;
		CloseHandle(COM);
		if (numDevice == 1)
		{
			message = "Error: �� ������� ������� ������ � ������� ���������� ����������� �������� ��� "
				"����������������� ���������� (���������� ����� 1). \n��. ���� logfile.txt ��� ��������� ����������.";
			MessageBox(NULL, message.c_str(), "Error", MB_OK | MB_ICONERROR);
			Date(date, sizeof(date));
			message = ": >������: �� ������� ������� ������ � ������� ���������� ����������� �������� "
				"��� ���������� ����������������� ����������. \n������ ������ ������ ���������� "
				"��� ���������� � ����� ���������� ����������� ��� ��� ���������� � ������������ ��������.";
			logFile << date << " ����������" << numDevice << message << endl;
			message = "���������� ������������� ��, ����� ���������� ��������� � ����� ���������� ����������� ����������. "
				"���� ������ �������� �� �������, ���������� �������� ����������� ����������.";
			logFile << message << endl;
			Date(date, sizeof(date));
			logFile << date << ": >��������� ����������� � �������.";
			logFile.close();
			exit(0);
		}
		if (numDevice == 2)
		{
			message = "Error: �� ������� ������� ������ � ������� ���������� ����������� �������� ��� "
				"����������������� ���������� (���������� ����� 2). \n��. ���� logfile.txt ��� ��������� ����������.";
			MessageBox(NULL, message.c_str(), "Warning", MB_OK | MB_ICONWARNING);
			Date(date, sizeof(date));
			message = ": >������: �� ������� ������� ������ � ������� ���������� ����������� �������� "
				"��� ���������� ����������������� ����������. \n������ ������ ������ ���������� ��� ���������� "
				"� ����� ���������� ����������� ��� ��� ���������� � ������������ ��������.";
			logFile << date << " ����������" << numDevice << message << endl;
			message = "���������� ������������� ��, ����� ���������� ��������� � ����� ���������� "
				"����������� ����������. ���� ������ �������� �� �������, ���������� �������� ����������� ����������.";
			logFile << message << endl;
			return;
		}
	}
	Date(date, sizeof(date));
	message = ": >���������� ������ � ������� ���������� ����������� �������� ��� ����������������� ���������� ��������� �������.";
	logFile << date << " ����������" << numDevice << message << endl;
	
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
			message = "Error: ������� ������������ ��������� ���������� ���������������� ����������� ��� ��������� "
				"����� �� ������ (���������� ����� 1). \n��. ���� logfile.txt ��� ��������� ����������.";
			MessageBox(NULL, message.c_str(), "Error", MB_OK | MB_ICONERROR);
			Date(date, sizeof(date));
			message = ": >������: ������� ������������ �������� ��������� DCB "
				"(��������� ���������� ���������������� �����������) ��� �������� ����� �� ������.";
			logFile << date << " ����������" << numDevice << message << endl;
			message = "������ ������ ���������� ��� ������������ ��������� ������������ ����������. "
				"��������� ������������ �� ����� conf.ini "
				"\n������������� ���������: baudrate=9600, bytesize=8, parity=3(NOPARITY), stopbits=1(ONESTOPBIT).";
			logFile << message << endl;
			Date(date, sizeof(date));
			logFile << date << ": >��������� ����������� � �������.";
			logFile.close();
			exit(0);
		}
		if (numDevice == 2)
		{
			message = "Error: ������� ������������ ��������� ���������� ���������������� ����������� ��� "
				"��������� ����� �� ������ (���������� ����� 2). \n��. ���� logfile.txt ��� ��������� ����������.";
			MessageBox(NULL, message.c_str(), "Warning", MB_OK | MB_ICONWARNING);
			Date(date, sizeof(date));
			message = ": >������: ������� ������������ �������� ��������� DCB "
				"(��������� ���������� ���������������� �����������) ��� �������� ����� �� ������.";
			logFile << date << " ����������" << numDevice << message << endl;
			message = "������ ������ ���������� ��� ������������ ��������� ������������ ����������. "
				"��������� ������������ �� ����� conf.ini "
				"\n������������� ���������: baudrate2=9600, bytesize2=8, parity2=3(NOPARITY), stopbits2=1(ONESTOPBIT).";
			logFile << message << endl;
			return;
		}
 	}
	Date(date, sizeof(date));
	message = ": >��������� �������� �������� ���������� ���������������� ����������� ��������� �������.";
	logFile << date << " ����������" << numDevice << message << endl;

	if(!SetCommMask(COM, EV_RXCHAR)) 
	{                                   
		COM = INVALID_HANDLE_VALUE;
		CloseHandle(COM);
		if (numDevice == 1)
		{
			message = "Error: ��������� ������ �������, ������� ����� ������������� ��� ������� "
				"����������������� ���������� ���������� (���������� ����� 1). \n��. ���� logfile.txt ��� ��������� ����������.";
			MessageBox(NULL, message.c_str(), "Error", MB_OK | MB_ICONERROR);
			Date(date, sizeof(date));
			message = ": >������: ��������� ����� EV_RXCHAR ��� ������� ����������������� ���������� ����������.";
			logFile << date << " ����������" << numDevice << message << endl;
			message = "������ ��� ���������� ������� SetCommMask(COMPort, EV_RXCHAR) ��� ������������ ������������ "
				"������� �� ����������������� ����������. ����� EV_RXCHAR: ������ ��� ������ � ������� � ����� ����� ������. "
				"\n���������� �������� ��������� ������������ � ������������� ���. ���� �� �������� - ���������� �������� "
				"����������� ����������.";
			logFile << message << endl;
			Date(date, sizeof(date));
			logFile << date << ": >��������� ����������� � �������.";
			logFile.close();
			exit(0);
		}
		if (numDevice == 2)
		{
			message = "Error: ��������� ������ �������, ������� ����� ������������� ��� ������� "
				"����������������� ���������� ���������� (���������� ����� 2). \n��. ���� logfile.txt ��� ��������� ����������.";
			MessageBox(NULL, message.c_str(), "Warning", MB_OK | MB_ICONWARNING);
			Date(date, sizeof(date));
			message = ": >������: ��������� ����� EV_RXCHAR ��� ������� ����������������� ���������� ����������.";
			logFile << date << " ����������" << numDevice << message << endl;
			message = "������ ��� ���������� ������� SetCommMask(COMPort, EV_RXCHAR) ��� ������������ "
				"������������ ������� �� ����������������� ����������. ����� EV_RXCHAR: ������ ��� ������ "
				"� ������� � ����� ����� ������. \n���������� �������� ��������� ������������ � ������������� ���. "
				"���� �� �������� - ���������� �������� ����������� ����������.";
			logFile << message << endl;
			return;
		}
 	}
	Date(date, sizeof(date));
	message = ": >��������� ������ �������, ������� ����� ������������� ��� ������� ����������������� ����������, ��������� �������.";
	logFile << date << " ����������" << numDevice << message << endl;
	
 	ct.ReadIntervalTimeout = MAXDWORD; 
 	ct.ReadTotalTimeoutMultiplier = 0; 
 	ct.ReadTotalTimeoutConstant = 0; 
	
 	if(!SetCommTimeouts(COM, &ct)) 
	{
		COM = INVALID_HANDLE_VALUE;
		CloseHandle(COM);
		if (numDevice == 1)
		{
			message = "Error: ������� ������������ ��������� ��������� ������� �� ������� ��� ����������������� "
				"���������� ��� ��������� ����� �� ������ (���������� ����� 1). "
				"\n��. ���� logfile.txt ��� ��������� ����������.";
			MessageBox(NULL, message.c_str(), "Error", MB_OK | MB_ICONERROR);
			Date(date, sizeof(date));
			message = ": >������: ������� ������������ �������� ��������� COMMTIMEOUTS "
				"(��������� ��������� ������� �� ������� ��� ����������������� ����������) ��� �������� �� ������.";
			logFile << date << " ����������" << numDevice << message << endl;
			message = "���������� �������� ��������� ������������ � ������������� ���. ���� �� "
				"�������� - ���������� �������� ����������� ����������.";
			logFile << message << endl;
			Date(date, sizeof(date));
			logFile << date << ": >��������� ����������� � �������.";
			logFile.close();
			exit(0);
		}
		if (numDevice == 2)
		{
			message = "Error: ������� ������������ ��������� ��������� ������� �� ������� ��� ����������������� "
				"���������� ��� ��������� ����� �� ������ (���������� ����� 2). "
				"\n��. ���� logfile.txt ��� ��������� ����������.";
			MessageBox(NULL, message.c_str(), "Warning", MB_OK | MB_ICONWARNING);
			Date(date, sizeof(date));
			message = ": >������: ������� ������������ �������� ��������� COMMTIMEOUTS "
				"(��������� ��������� ������� �� ������� ��� ����������������� ����������) ��� �������� �� ������.";
			logFile << date << " ����������" << numDevice << message << endl;
			message = "���������� �������� ��������� ������������ � ������������� ���. "
				"���� �� �������� - ���������� �������� ����������� ����������.";
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
	message = ": >��������� ���������� ��������� ������� �� ������� ��� ����������������� ���������� ��������� �������.";
	logFile << date << " ����������" << numDevice << message << endl;
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
		cell->SetWString(L"���");
		cell = sheet->Cell(0, 2);
		cell->SetWString(L"����");
		cell = sheet->Cell(0, 3);
		cell->SetWString(L"�����");
		cell = sheet->Cell(0, 4);
		cell->SetWString(L"���������");
		cell = sheet->Cell(0, 5);
		cell->SetWString(L"������");
		cell = sheet->Cell(0, 6);
		cell->SetWString(L"������������� (-�)");
		cell = sheet->Cell(0, 7);
		cell->SetWString(L"����������");
		cell = sheet->Cell(0, 8);
		cell->SetWString(L"��� �������");
		cell = sheet->Cell(0, 9);
		cell->SetWString(L"������");

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
		logFile << date << ": >�������� Excel-����� ��������� �������." << endl;
	}
	xlsCreate = false;
	
	if (pointer->numDevice == 1)
	{
		Date(date, sizeof(date));
		logFile << date << " ����������1" << ": >������������� " << readId << " ��� ������� ������� � Excel-����." << endl;
	}
	if (pointer->numDevice == 2)
	{
		char* format = "%d.%m.%Y %H:%M:%S";
		Date(date, sizeof(date));
		logFile << date << " ����������2" << ": >������������� " << readId << " ��� ������� ������� � Excel-����." << endl;
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
		message = "Error: ��� �������� ����������� ������� ��� ����������� �������� �������� ������. "
			"\n��. ���� logfile.txt ��� ��������� ����������.";
		MessageBox(NULL, message.c_str(), "Error", MB_OK | MB_ICONERROR);
		Date(date, sizeof(date));
		message = ": >��� �������� ����������� ������� ��� ����������� �������� �������� ������.";
		logFile << date << message << endl;
		message = "���������� ������������� ��, ����� ���������� ��������� � ����� ���������� ����������� ����������. "
			"���� ������ �������� �� �������, ���������� �������� ����������� ����������.";
		logFile << message << endl;
		Date(date, sizeof(date));
		logFile << date << ": >��������� ����������� � �������.";
		logFile.close();
		exit(0);
	}

	while(1)				   
	{
		if (!WaitCommEvent(COM, &mask, &overlapped)) 
		{
			if (!GetLastError() == ERROR_IO_PENDING)
			{
				message = "Error: ������� WaitCommEvent ����������� � �������. "
					"\n��. ���� logfile.txt ��� ��������� ����������.";
				MessageBox(NULL, message.c_str(), "Error", MB_OK | MB_ICONERROR);
				Date(date, sizeof(date));
				logFile << date << ": >������� WaitCommEvent ����������� � �������." << endl;
				message = "���������� ������������� ��, ����� ���������� ��������� � ����� ���������� ����������� ����������. "
						"���� ������ �������� �� �������, ���������� �������� ����������� ����������.";
				logFile << message << endl;
				Date(date, sizeof(date));
				logFile << date << ": >��������� ����������� � �������.";
				logFile.close();
				exit(0);
			}
		}

		signal = WaitForSingleObject(overlapped.hEvent, INFINITE);	
		if(signal == WAIT_OBJECT_0)				        
		{
			if(!GetOverlappedResult(COM, &overlapped, &temp, true))
			{
				message = "Error: ��������� ������ � �������� ����������. \n��. ���� logfile.txt ��� ��������� ����������.";
				MessageBox(NULL, message.c_str(), "Error", MB_OK | MB_ICONERROR);
				Date(date, sizeof(date));
				message = ": >��������� ������ � �������� ����������, ���� ���������� ���� ��������� �� ����� ������ ���������.";
				logFile << date << message << endl;
				message = "������ ������ �������� ���������� ���������� �� ����� ������ ���������. "
					"���������� ������������� ��, ����� ���������� ��������� � ����� ���������� ����������� ����������. "
					"���� ������ �������� �� �������, ���������� �������� ����������� ����������.";
				logFile << message << endl;
				Date(date, sizeof(date));
				logFile << date << ": >��������� ����������� � �������.";
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