// Copyright(c) 2016 YamaArashi
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE.

#include <cstdio>
#include <cstdlib>
#include <cstring>
#include <cctype>
#include <cassert>
#include <string>
#include <set>
#include <iostream>
#include "msoxfix.h"
#include "error.h"

//FILE* g_inputFile = nullptr;
//FILE* g_outputFile = nullptr;

[[noreturn]] static void PrintUsage()
{
    std::printf(
        "Usage: MSOXFIX filename.msox\n"
        "\n"
        "    input_file  filename(.MSOX) of CHK file\n"
        "   output_file  input_file(.Microsoft Office file extension)\n"
        "\n"
        "Jose Luis Rincon Ciriano (c) 2021\n"
    );
    std::exit(1);
}

static std::string StripExtension(std::string s)
{
    std::size_t pos = s.find_last_of('.');

    if (pos > 0 && pos != std::string::npos)
    {
        s = s.substr(0, pos);
    }

    return s;
}

static std::string GetExtension(std::string s)
{
    std::size_t pos = s.find_last_of('.');

    if (pos > 0 && pos != std::string::npos)
    {
        return s.substr(pos + 1);
    }

    return "";
}


int main(int argc, char** argv)
{
    std::string inputFilename;
    std::string outputFilename;

    for (int i = 1; i < argc; i++)
    {
        if (inputFilename.empty())
            inputFilename = argv[i];
        else if (outputFilename.empty())
            outputFilename = argv[i];
        else
            PrintUsage();
    }

    if (inputFilename.empty())
        PrintUsage();

    if (GetExtension(inputFilename) != "msox")
        RaiseError("input filename extension is not \"msox\"");

    //if (GetExtension(outputFilename) != "s")
    //    RaiseError("output filename extension is not \"s\"");

    if (outputFilename.empty())
        outputFilename = StripExtension(inputFilename);

    switch (msoxfix(inputFilename, outputFilename))
    {
    case 1:
        RaiseError("failed to open \"%s\" for reading", inputFilename.c_str());
        break;
    case 2:
        RaiseError("[Content_Types] .xmlPK not found\nThe file will not be renamed\n");
        break;
    case 3:
        RaiseError("The file could not be renamed\nbecause the type could not be defined.");
        break;
    case 0:
    default:
        std::wcout << "Correct file extension successfully\n";
        break;
    }

    //std::fclose(g_inputFile);
    //std::fclose(g_outputFile);
    return 0;
}
