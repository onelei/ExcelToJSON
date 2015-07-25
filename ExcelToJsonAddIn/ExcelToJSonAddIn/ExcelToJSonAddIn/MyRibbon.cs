using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
namespace ExcelToJSonAddIn
{
    public partial class MyRibbon
    {
        private int KeyRow = 2;
        private int KeyColumn = 1;
        private int ValueRow = 4;
        private int ValueColumn = 1;
        private void MyRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            Key_R.Text = "" + KeyRow;
            Key_C.Text = "" + KeyColumn;
            Value_R.Text = "" + ValueRow;
            Value_C.Text = "" + ValueColumn;
        }

        private void button_JSon_Click(object sender, RibbonControlEventArgs e)
        {       
            // When click the button, start switch;
            ExcelToJSON();
        }

        /**
         * ExcelToJSON
         * @function swich Excel to JSON; 
         * return 
         */
        private void ExcelToJSON()
        {
            CheckInputValue();
            Stopwatch MyCodeExeTime = new Stopwatch();
            MyCodeExeTime.Start();
            string text = "";
            int totalRows = ThisAddIn.Instance.workBook.ActiveSheet.UsedRange.Rows.Count;
            int totalColumns = ThisAddIn.Instance.workBook.ActiveSheet.UsedRange.Columns.Count;
            totalRows=FixtotalRows( totalRows);
            totalColumns = FixtotalColumns(totalColumns);

            text += "[{";
            for (int i = ValueRow; i != totalRows + 1; ++i)
            {
                for (int j = ValueColumn; j != totalColumns + 1; ++j)
                {               
                       // key =>    "ID:"   ;
                        text += "\"" + ThisAddIn.Instance.workBook.ActiveSheet.Cells(KeyRow, j).Value + "\":";
                        string type = ""+ThisAddIn.Instance.workBook.ActiveSheet.Cells(KeyRow+1, j).Value ;
                        bool isIntValue  =  JugeIsIntType( type);
                        // In the last row and last column, modify text style;
                        if (i == totalRows && j == totalColumns)
                        {
                            // value =>  "value"  ;
                            string value = ""+ThisAddIn.Instance.workBook.ActiveSheet.Cells(i, j).Value ;
                            text += FixValueAdd(value, type);
                        }
                        else
                        {
                            // Every Row Start, write "},";
                            if (j == totalColumns)
                            {
                                string value = "" + ThisAddIn.Instance.workBook.ActiveSheet.Cells(i, j).Value;
                                text += FixValueAdd(value,type) +"},{";
                            }
                            else
                            {
                                string value = "" + ThisAddIn.Instance.workBook.ActiveSheet.Cells(i, j).Value;
                                text +=  FixValueAdd(value,type) +",";
                            }
                        }
                    
                }
            }
            text += "}]";

            // Write to file;
            string path = ThisAddIn.Instance.workBook.Path+"\\";

            string fileName = ThisAddIn.Instance.workBook.Name;
            fileName = fileName.Replace(".xlsx", "");
            fileName += ".json";      
            WriteToFile(path, fileName, text);
            MyCodeExeTime.Stop();
            string myTime = MyCodeExeTime.ElapsedMilliseconds.ToString();
            ShowMessage(totalRows, totalColumns,myTime);
        }

        /// <summary>
        /// write file;
        /// </summary>
        /// <param name="path"></param>
        /// <param name="name"></param>
        /// <param name="text"></param>
        public void WriteToFile(string path, string name, string text)
        {
            lock (this)
            {
                // File stream information;
                StreamWriter sw;
                FileInfo t = new FileInfo(path + name);
                // Create File Stream;
                sw = t.CreateText();
                // Write text by line style;
                sw.WriteLine(text);
                // Close stream;
                sw.Close();
                // Destrory stream;
                sw.Dispose();
            }

        }

        /*
         * Show message.
         */
        private void ShowMessage(int totalRows, int totalColumns,string _ExeTime)
        {
            MessageBox.Show(  
                "\n"
                + "  Excel to JSON\n"
                + "  Save Successful!\n\n"
                + "  Time:  " + _ExeTime + "毫秒.\n"
                + "  Total: " + totalRows + " Rows"
                + ", " + totalColumns + " Columns");
        }

        private bool JugeNumberOrNot(string _text)
        {
            if (String.IsNullOrEmpty(_text))
            {
                MessageBox.Show("Input value is null or empty");
                return false;
            }
            else
            {
                 try
                 {
                     Int32.Parse(_text);
                 }
                 catch
                 {
                     MessageBox.Show("Input value is not number!");
                     return false;
                 }
            }
            return true;
        }

        private void ShowAbout()
        {
            MessageBox.Show(
                  "\n"
                + "  Excel to JSON\n"
                + "  Created by OneLei.\n"
                + "  Email: ahleiwolong@163.com \n"
                + "  Copyright (c) 2015 Year. All rights reserved.\n\n");
        }

        private void button_About_Click(object sender, RibbonControlEventArgs e)
        {
            ShowAbout();
        }

        /// <summary>
        /// check the input value , make sure value is integer.
        /// </summary>
        void CheckInputValue()
        {
            // Get the key;
            if (JugeNumberOrNot(Key_R.Text))
            {
                KeyRow = System.Int32.Parse(Key_R.Text);
            }
            if (JugeNumberOrNot(Key_C.Text))
            {
                KeyColumn = System.Int32.Parse(Key_C.Text);
            }

            // Get the value;
            if (JugeNumberOrNot(Value_R.Text))
            {
                ValueRow = System.Int32.Parse(Value_R.Text);
            }
            if (JugeNumberOrNot(Value_C.Text))
            {
                ValueColumn = System.Int32.Parse(Value_C.Text);
            }
        }

        private int FixtotalRows(int totalRows)
        {
            for (int i = 1; i < totalRows;i++ )
            {
               string cellValue = ""+ThisAddIn.Instance.workBook.ActiveSheet.Cells(i, 1).Value;
                if (string.IsNullOrEmpty(cellValue))
                {
                    // 当前的为空,则totalRow到此为止,修复totalRow;
                    return i-1;
                }
            }
            return totalRows;
        }

        private int FixtotalColumns(int totalColumns)
        {
            for (int i = 1; i < totalColumns; i++)
            {
                string cellValue = ""+ThisAddIn.Instance.workBook.ActiveSheet.Cells(1, i).Value;
                if (string.IsNullOrEmpty(cellValue))
                {
                    // 当前的为空,则totalRow到此为止,修复totalRow;
                    return i-1;
                }
            }
            return totalColumns;
        }

        bool JugeIsStringType(string type)
        {
            if(type=="string"||type=="String")
            {
                return true;
            }
            return false;
        }

        bool JugeIsIntType(string type)
        {
            if(type=="int"||type=="Int")
            {
                return true;
            }
            return false;
        }

        string FixValueAdd(string value,string type)
        {
            if(JugeIsIntType(type))
            {
                return value;
            }
            return "\"" + value + "\"";
        }

    }
}
