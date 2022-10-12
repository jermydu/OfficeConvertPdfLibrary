// TestOfficeConvert.cpp : 此文件包含 "main" 函数。程序执行将在此处开始并结束。
//

#include <iostream>

#using "../x64/Debug/OfficeConvertPdfLibrary.dll"
using namespace OfficeConvertPdfLibrary;

int main()
{
	ClassOfficeConvertPdfLibrary^ pConvertClass = gcnew ClassOfficeConvertPdfLibrary();
	//路径一定要注意  错误示范
	//D:\sourcecode\OfficeConvertPdfLibrary\x64\Debug\test.xlsx
	//D:/sourcecode/OfficeConvertPdfLibrary/x64/Debug/test.xlsx
	int nlsxResult = pConvertClass->ExcelConvertPdf("D:\\sourcecode\\OfficeConvertPdfLibrary\\x64\\Debug\\test.xlsx","D:\\sourcecode\\OfficeConvertPdfLibrary\\x64\\Debug\\test_xlsx.pdf");
	if (nlsxResult == 0)
	{
		std::cout << "test.xlsx 转换失败" << std::endl;
	}
	else
	{
		std::cout << "test.xlsx 转换成功" << std::endl;
	}
	int wordResult = pConvertClass->WordConvertPdf("D:\\sourcecode\\OfficeConvertPdfLibrary\\x64\\Debug\\test.docx","D:\\sourcecode\\OfficeConvertPdfLibrary\\x64\\Debug\\test_word.pdf");
	if (wordResult == 0)
	{
		std::cout << "test.docx 转换失败" << std::endl;
	}
	else
	{
		std::cout << "test.docx 转换成功" << std::endl;
	}

	int pptResult = pConvertClass->PowerPointConvertPdf("D:\\sourcecode\\OfficeConvertPdfLibrary\\x64\\Debug\\test.pptx","D:\\sourcecode\\OfficeConvertPdfLibrary\\x64\\Debug\\test_ppt.pdf");
	if (pptResult == 0)
	{
		std::cout << "test.pptx 转换失败" << std::endl;
	}
	else
	{
		std::cout << "test.pptx 转换成功" << std::endl;
	}
	
}

