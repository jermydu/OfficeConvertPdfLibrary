// TestOfficeConvertCpp.cpp : 此文件包含 "main" 函数。程序执行将在此处开始并结束。
//

#include <iostream>
#include "../CppOfficeConvertPdf/CppOfficeConvertPdf.h"

#pragma comment (lib,"../x64/Debug/CppOfficeConvertPdf.lib")   

int main()
{
    CCppOfficeConvertPdf *pConvert = new CCppOfficeConvertPdf();
    pConvert->ConvertPdf("xlsx", "D:\\sourcecode\\OfficeConvertPdfLibrary\\x64\\Debug\\test.xlsx", "D:\\sourcecode\\OfficeConvertPdfLibrary\\x64\\Debug\\cpp_xlsx.pdf");
    std::cout << "xlsx 转换完成" << std::endl;
    pConvert->ConvertPdf("docx", "D:\\sourcecode\\OfficeConvertPdfLibrary\\x64\\Debug\\test.docx","D:\\sourcecode\\OfficeConvertPdfLibrary\\x64\\Debug\\cpp_word.pdf");
    std::cout << "docx 转换完成" << std::endl;
    pConvert->ConvertPdf("pptx", "D:\\sourcecode\\OfficeConvertPdfLibrary\\x64\\Debug\\test.pptx", "D:\\sourcecode\\OfficeConvertPdfLibrary\\x64\\Debug\\cpp_ppt.pdf");
    std::cout << "pptx 转换完成" << std::endl;
}
