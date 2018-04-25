#include "Copy.h"

bool CheckLang(int n, char* str)
{
	for (int i = 0; i < n; i++)
	{
		if (!(str[i] >= 32 && str[i] <= 127))
		{
			return false;
		}
	}
	return true;
}

void InitOpenFile(OPENFILENAME &ofn, char *szFile, int n)
{
	ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = NULL;
    ofn.lpstrDefExt = 0;
    ofn.lpstrFile = szFile;
    ofn.lpstrFile[0] = '\0';
    ofn.nMaxFile = n;
    ofn.lpstrFilter = "Excel files (.xls)\0*.XLS\0";
    ofn.nFilterIndex = 1;
    ofn.lpstrInitialDir = 0;
    ofn.lpstrTitle = 0;
	ofn.lpstrFileTitle = NULL;
    ofn.nMaxFileTitle = 0;
    ofn.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST | OFN_ALLOWMULTISELECT | OFN_EXPLORER;
}

void CopyInf(char *fileIn, char *fileOut)
{
	BasicExcel xls(fileIn);
	BasicExcelWorksheet* sheet = xls.GetWorksheet(0);
	BasicExcelCell* cell;

	BasicExcel xlsOut(fileOut);
	BasicExcelWorksheet* sheetOut = xlsOut.GetWorksheet(0);
	BasicExcelCell* cellOut;
	XLSFormatManager fmtMgr(xlsOut);
	ExcelFont fontBold;
	fontBold._weight = FW_BOLD;
	CellFormat fmtBold(fmtMgr);
	fmtBold.set_font(fontBold);

	size_t maxRows = sheet->GetTotalRows();
	size_t maxCols = sheet->GetTotalCols();
	size_t maxRowsOut = sheetOut->GetTotalRows();
	size_t maxColsOut = sheetOut->GetTotalCols();
	size_t col = 0;
	size_t row = 0;
	size_t colOut = 0;
	size_t rowOut = 0;

	string buf;
	string bufOut;
	wstring wBuf;
	wstring wBufOut;
	int number;
	
	cellOut = sheetOut->Cell(0, 0);
	buf = cellOut->Type();
	if ((maxRowsOut == 1) && (maxColsOut == 1) && (buf.front() == 0))
	{
		sheetOut->SetColWidth(1, 3000);
		sheetOut->SetColWidth(2, 5000);
		sheetOut->SetColWidth(3, 4000);

		cell = sheet->Cell(0, 9);
		wBuf = cell->GetWString();
		cellOut = sheetOut->Cell(0, 0);
		cellOut->SetWString(wBuf.c_str());
				
		cell = sheet->Cell(1, 9);
		wBuf = cell->GetWString();
		cellOut = sheetOut->Cell(1, 0);
		cellOut->SetWString(wBuf.c_str());

		cell = sheet->Cell(0, 6);
		wBuf = cell->GetWString();
		cellOut = sheetOut->Cell(0, 2);
		cellOut->SetWString(wBuf.c_str());

		cell = sheet->Cell(1, 6);
		wBuf = cell->GetWString();
		cellOut = sheetOut->Cell(1, 2);
		cellOut->SetWString(wBuf.c_str());

		cellOut = sheetOut->Cell(6, 0);
		cellOut->SetWString(L"№");
		cellOut = sheetOut->Cell(6, 1);
		cellOut->SetString("ID");
		cellOut = sheetOut->Cell(6, 2);
		cellOut->SetWString(L"ФИО");
				
		for (row = 1 ; row < maxRows; row++)
		{
			cell = sheet->Cell(row, col);
			buf = cell->GetString();
			cellOut = sheetOut->Cell(row+6, 1);
			cellOut->SetString(buf.c_str());
			cellOut = sheetOut->Cell(row+6, 0);
			cellOut->Set((int)row);
			cellOut = sheetOut->Cell(row+6, 3);
			cellOut->Set("+");
			buf="";
		}
				
		cell = sheet->Cell(1, 2);
		buf = cell->GetString();
		cellOut = sheetOut->Cell(6, 3);
		cellOut->SetString(buf.c_str());
				
		cell = sheet->Cell(1, 8);
		wBuf = cell->GetWString();
		cellOut = sheetOut->Cell(5, 3);
		cellOut->SetWString(wBuf.c_str());

		cellOut = sheetOut->Cell(0, 0);
		cellOut->SetFormat(fmtBold);
		cellOut = sheetOut->Cell(0, 2);
		cellOut->SetFormat(fmtBold);
		cellOut = sheetOut->Cell(6, 0);
		cellOut->SetFormat(fmtBold);
		cellOut = sheetOut->Cell(6, 1);
		cellOut->SetFormat(fmtBold);
		cellOut = sheetOut->Cell(6, 2);
		cellOut->SetFormat(fmtBold);
	}
	else
	{
		cell = sheet->Cell(1, 6);
		wBuf = cell->GetWString();
		cellOut = sheetOut->Cell(1, 2);
		wBufOut = cellOut->GetWString();
		rowOut = 2;
		colOut = 2;
		if (!(wBuf == wBufOut))
		{
			cellOut = sheetOut->Cell(rowOut, colOut);
			buf = cellOut->Type();
			if (buf.front() == 0)
				cellOut->SetWString(wBuf.c_str());
			else
			{
				wBufOut = cellOut->GetWString();
				if (!(wBuf == wBufOut))
				{
					rowOut++;
					cellOut = sheetOut->Cell(rowOut, colOut);
					buf = cellOut->Type();
					if (buf.front() == 0)
					{
						cellOut = sheetOut->Cell(rowOut, colOut);
						cellOut->SetWString(wBuf.c_str());
					}
				}
			}
		}

		cell = sheet->Cell(1, 2);
		buf = cell->GetString();
		cellOut = sheetOut->Cell(6, maxColsOut);
		cellOut->SetString(buf.c_str());
		sheetOut->SetColWidth(maxColsOut, 6000);

		cell = sheet->Cell(1, 8);
		wBuf = cell->GetWString();
		cellOut = sheetOut->Cell(5, maxColsOut);
		cellOut->SetWString(wBuf.c_str());

		for (row = 1 ; row < maxRows; row++)
		{
			bool idFound = false;
			bool endCheck = false;

			col = 0;
			cell = sheet->Cell(row, col);
			buf = cell->GetString();
				
			for (rowOut = 7 ; rowOut < maxRowsOut; rowOut++)
			{
				colOut = 1;
				cellOut = sheetOut->Cell(rowOut, colOut);
				bufOut = cellOut->GetString();
				if (buf == bufOut)
				{
					cellOut = sheetOut->Cell(rowOut, maxColsOut);
					cellOut->SetString("+");
					idFound = true;
					break;
				}

				if ((maxRowsOut - rowOut) == 1)
					endCheck = true;

				if ((idFound == false) && (endCheck == true))
				{
					cellOut = sheetOut->Cell(maxRowsOut, colOut);
					cellOut->SetString(buf.c_str());
					cellOut = sheetOut->Cell(maxRowsOut, maxColsOut);
					cellOut->SetString("+");
					cellOut = sheetOut->Cell(maxRowsOut, 0);
					number = maxRowsOut - 6;
					cellOut->Set((int)(maxRowsOut - 6));
					maxRowsOut = maxRowsOut + 1;
				}
			}
		}

		for (colOut = 3; colOut <= maxColsOut; colOut++)
		{
			for (rowOut = 7; rowOut < maxRowsOut; rowOut++)
			{
				cellOut = sheetOut->Cell(rowOut, colOut);
				if (cellOut->Type() == 0)
					cellOut->SetString("-");
			}
		}
	}
	xlsOut.Save();
}

void SelectFileTo(char *filesIn, int size, bool multipleFiles)
{
	OPENFILENAME ofnSecond;
	string message;
	char fileOut[256] = {0};

	ZeroMemory(&ofnSecond, sizeof(ofnSecond));
	InitOpenFile(ofnSecond, fileOut, sizeof(fileOut));
	ofnSecond.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST;

	message = "Выберите Excel-файл, в который необходимо скопировать информацию. "
			"\nВажно: Excel-файл должнен быть закрыт и путь к файлу не должен содержать русских символов.";
	MessageBox(NULL, message.c_str(), "Copy", MB_OK | MB_ICONINFORMATION);

	if (GetOpenFileName(&ofnSecond))
	{
		if (CheckLang(strlen(fileOut), fileOut) == false)
		{
			MessageBox(0, "Error: Путь к Excel-файлу содержит русские символы.", "Copy", MB_OK | MB_ICONERROR);
			exit(0);
		}
		if (!multipleFiles)
		{
			if (!(memcmp(fileOut, filesIn, sizeof(fileOut))))
			{
				MessageBox(0, "Error: Выбран один и тот же файл. Пожалуйста, повторите попытку.", "Copy", MB_OK | MB_ICONERROR);
				exit(0);
			}

			CopyInf(filesIn, fileOut);

			message = "Информация из файла " + (string)filesIn + " в файл " + (string)fileOut + " успешно скопирована.";
			MessageBox(0, message.c_str(), "Copy", MB_OK | MB_ICONINFORMATION);
		}
		else
		{
			char dirFile[256] = {0};
			char dirFolder[256] = {0};

			for (int i = 0; i < size; i++)
			{
				if ((filesIn[i] == '\0') && (filesIn[i+1] == '\0'))
					break;
				if (filesIn[i] == '\0')
					filesIn[i] = '@';
			}

			char *pCh = strtok(filesIn, "@");
			if (strlen(pCh) != 3)
			{
				strcat(dirFolder, pCh);
				strcat(dirFolder, "\\");
			}
			else
				strcat(dirFolder, pCh);

			pCh = strtok(NULL, "@");
			while (pCh != NULL)
			{
				strcat(dirFile, dirFolder);
				strcat(dirFile, pCh);

				CopyInf(dirFile, fileOut);

				message = "Информация из файла " + (string)dirFile + " в файл " + (string)fileOut + " успешно скопирована.";
				MessageBox(0, message.c_str(), "Copy", MB_OK | MB_ICONINFORMATION);

				pCh = strtok(NULL, "@");
				memset(dirFile, 0, sizeof(dirFile));
			}
		}
	}
	else
	{	
		MessageBox(0, "Программа завершена клавишей \"Отмена\". \nКопирование отменено.", "Copy", MB_OK | MB_ICONINFORMATION);
		exit(0);
	}
}

void SelectFilesFrom()
{
	OPENFILENAME ofnFirst;
	string message;
	char filesIn[1024] = {0};
	bool multipleFiles;

	ZeroMemory(&ofnFirst, sizeof(ofnFirst));
	InitOpenFile(ofnFirst, filesIn, sizeof(filesIn));

	message = "Выберите Excel-файлы, из которых необходимо скопировать информацию. "
			  "\nВажно: Excel-файлы должны быть закрыты и пути к данным файлам не должны содержать русских символов.";
					
	MessageBox(NULL, message.c_str(), "Copy", MB_OK | MB_ICONINFORMATION);
	if (GetOpenFileName(&ofnFirst))
	{
		int nOffset = ofnFirst.nFileOffset;
		if (nOffset > strlen(ofnFirst.lpstrFile))
		{
			if (CheckLang(strlen(filesIn), filesIn) == false)
			{
				message = "Error: Путь к одному из выбранных Excel-файлов содержит русские символы.";
				MessageBox(NULL, message.c_str(), "Copy", MB_OK | MB_ICONERROR);
				exit(0);
			}
			
			multipleFiles = true;
			SelectFileTo(filesIn, sizeof(filesIn), multipleFiles);
		}
		else
		{
			if (CheckLang(strlen(filesIn), filesIn) == false)
			{
				message = "Error: Путь к Excel-файлу содержит русские символы.";
				MessageBox(NULL, message.c_str(), "Copy", MB_OK | MB_ICONERROR);
				exit(0);
			}
			multipleFiles = false;
			SelectFileTo(filesIn, sizeof(filesIn), multipleFiles);
		}
	}
	else
	{	
		MessageBox(0, "Программа завершена клавишей \"Отмена\", либо было выбрано слишком большое количество "
			"файлов для копирования (рекомендуемое значение 5). \nКопирование отменено.", "Copy", MB_OK | MB_ICONINFORMATION);
		exit(0);
	}
}