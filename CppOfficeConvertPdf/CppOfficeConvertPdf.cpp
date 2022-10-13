// CppOfficeConvertPdf.cpp : 定义 DLL 的导出函数。
//

#include "pch.h"
#include "framework.h"
#include "CppOfficeConvertPdf.h"

#using "../x64/Debug/OfficeConvertPdfLibrary.dll"
using namespace OfficeConvertPdfLibrary;


// 这是导出变量的一个示例
CPPOFFICECONVERTPDF_API int nCppOfficeConvertPdf=0;

// 这是导出函数的一个示例。
CPPOFFICECONVERTPDF_API int fnCppOfficeConvertPdf(void)
{
    return 0;
}

// 这是已导出类的构造函数。
CCppOfficeConvertPdf::CCppOfficeConvertPdf()
{
    return;
}

bool CCppOfficeConvertPdf::ConvertPdf(std::string type, std::string inPath, std::string outPath,std::string pngPath)
{
	System::String^ strType = gcnew System::String(type.c_str());
	System::String^ strInPath = gcnew System::String(inPath.c_str());
	System::String^ strOutPath = gcnew System::String(outPath.c_str());
	System::String^ strPngPath = gcnew System::String(pngPath.c_str());
	ClassOfficeConvertPdfLibrary^ pConvertClass = gcnew ClassOfficeConvertPdfLibrary();
	if (strType == "xlsx")
	{
		int nlsxResult = pConvertClass->ExcelConvertPdf(strInPath, strOutPath);
		if (nlsxResult == 0)
		{
			return false;
		}
		else
		{
			return true;
		}
	}
	else if(strType == "docx")
	{
		int wordResult = pConvertClass->WordConvertPdf(strInPath, strOutPath);
		if (wordResult == 0)
		{
			return false;
		}
		else
		{
			return true;
		}
	}else if (strType == "pptx")
	{
		int pptResult = pConvertClass->PowerPointConvertPdf(strInPath, strOutPath, strPngPath);
		if (pptResult == 0)
		{
			return false;
		}
		else
		{
			return true;
		}
	}
}
