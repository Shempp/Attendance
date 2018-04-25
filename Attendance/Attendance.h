#ifndef ATTENDANCE_H
#define ATTENDANCE_H

#include "ExcelFormat.h"
#include "resource.h"

using namespace ExcelFormat;

extern HWND hWnd;
extern CRITICAL_SECTION cs;
extern ofstream logFile;
extern char readBuffer[256];

struct ComConfData
{
	BasicExcel xls;
	char room[4];
	char building[2];
	wchar_t teacher[11];
	wchar_t discipline[31];
	wchar_t subject[21];
	wchar_t group[11];

	string comPort; 
	int baudRate; 
	int byteSize; 
	int parity; 
	int stopBits; 

	int secDevActive;
	string comPort2;
	int byteSize2; 
	int baudRate2;
	int parity2; 
	int stopBits2; 

	int numDevice;
	HANDLE COM;
	HANDLE COM2;
};

VOID Date(char* str , int size , const char* format = "%d.%m.%Y %H:%M:%S");

VOID ReadConfigFile(ComConfData *pointer);

VOID Disconnect(HANDLE &COMPort, int numDevice);

VOID Connect(ComConfData *pointer, int numDevice);

VOID WriteExcel(ComConfData *pointer);

DWORD WINAPI ReadCom(ComConfData *pointer);

VOID InitNotifyIconData(NOTIFYICONDATA &nid);

#endif ATTENDANCE_H