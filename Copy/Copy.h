#ifndef COPY_H
#define COPY_H

#include "ExcelFormat.h"
#include "resource.h"

using namespace ExcelFormat;
using namespace std;

bool CheckLang(int n, char* str);

void InitOpenFile(OPENFILENAME &ofn, char *szFile, int n);

void CopyInf(char *fileIn, char *fileOut);

void SelectFileTo(char *fileIn, int size, bool multipleFiles);

void SelectFilesFrom();

#endif COPY_H