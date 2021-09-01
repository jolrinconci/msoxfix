// Copyright(c) 2021 Jose Luis Rincon Ciriano
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

#include <iostream>
#include <cstdlib>
#include <fstream>
#include <cstdio>
#include <conio.h>
#include "main.h"
#include "error.h"

using namespace std;

#define NUM_LETTERS 19
#define CONTENT_LENGTH 21

int msoxfix(string old_file_name, string new_file_name)
{
    //int content_length = 21;
    
    int file_length;
    char * content_buffer = new char[CONTENT_LENGTH];
    int content_offset = 0;
    
    int counter_type_word = 0;
    int counter_type_excel = 0;
    int counter_type_power = 0;

    string content_string_string  = "[Content_Types].xmlPK";
    string file_type_word_string  = "document.xml.relsPK";  //---"word/_rels/document.xml.relsPK";
    string file_type_excel_string = "workbook.xml.relsPK";  //-----"xl/_rels/workbook.xml.relsPK";
    string file_type_power_string = "entation.xml.relsPK";  //"ppt/_rels/presentation.xml.relsPK";
    
    char * content_string  = new char[CONTENT_LENGTH];
    char * file_type_word  = new char[CONTENT_LENGTH];
    char * file_type_excel = new char[CONTENT_LENGTH];
    char * file_type_power = new char[CONTENT_LENGTH];

    char * new_file_name_ptr;
    char * old_file_name_ptr = &old_file_name[0];

    ifstream temp_file(old_file_name.c_str(), ifstream::binary);

    if ( temp_file.fail() )
    {
        return 1;
    }
    else
    {
        content_string = &content_string_string[0];
        file_type_word = &file_type_word_string[0];
        file_type_excel = &file_type_excel_string[0];
        file_type_power = &file_type_power_string[0];

        temp_file.seekg(0, temp_file.end);
        file_length = temp_file.tellg();
        temp_file.seekg(0, temp_file.beg);

        for (int i = 0; i < file_length; i++)
        {
            temp_file.seekg(i, temp_file.beg);
            temp_file.read(content_buffer, CONTENT_LENGTH);

            if (content_offset)
            {
                for (int k = 0; k < NUM_LETTERS; k++)
                {
                    if(content_buffer[k] == file_type_word[k] && k == counter_type_word)
                    {
                        counter_type_word++;
                    }
                    if(content_buffer[k] == file_type_excel[k] && k == counter_type_excel)
                    {
                        counter_type_excel++;
                    }
                    if(content_buffer[k] == file_type_power[k] && k == counter_type_power)
                    {
                        counter_type_power++;
                    }
                    if (counter_type_word != k+1 && counter_type_excel != k+1 && counter_type_power != k+1)
                    {
                        k = NUM_LETTERS;
                        counter_type_word = 0;
                        counter_type_excel = 0;
                        counter_type_power = 0;
                    }
                }

            }
            else
            {
                for (int j = 0; j < CONTENT_LENGTH; j++)
                {
                    if(content_buffer[j] == content_string[j])
                    {
                        content_offset++;
                    }
                    else
                    {
                        j = CONTENT_LENGTH;
                        content_offset = 0;
                    }
                }
            }

            if(counter_type_word == NUM_LETTERS)
            {
                wcout << "word found ... "<< endl;
                i = file_length;
            }
            else if(counter_type_excel == NUM_LETTERS)
            {
                wcout << "excel found ... "<< endl;
                i = file_length;
            }
            else if(counter_type_power == NUM_LETTERS)
            {
                wcout << "power point ound ..." << endl;
                i = file_length;
            }
        }

        temp_file.close();

        if (!content_offset)
        {
            return 2;
            //wcout << "Content_Types].xmlPK not found\n";
        }
        if ( counter_type_word )
        {
            new_file_name += ".docx";
            new_file_name_ptr = &new_file_name[0];
            rename(old_file_name_ptr,new_file_name_ptr);
        }
         else if ( counter_type_excel)
        {
            new_file_name += ".xlsx";
            new_file_name_ptr = &new_file_name[0];
            rename(old_file_name_ptr,new_file_name_ptr);
        }
         else if ( counter_type_power)
        {
            new_file_name += ".pptx";
            new_file_name_ptr = &new_file_name[0];
            rename(old_file_name_ptr,new_file_name_ptr);
        }
        else
        {
            return 3;
        }
    }
    return 0;
}
