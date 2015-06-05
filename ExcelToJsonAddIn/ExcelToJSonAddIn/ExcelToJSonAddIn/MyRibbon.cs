using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.IO;
namespace ExcelToJSonAddIn
{
    public partial class MyRibbon
    {
        private void MyRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnToJSon_Click(object sender, RibbonControlEventArgs e)
        {
            // When click the button, start switch;
            ExcelToJSON();  
        }


        /**
         * ExcelToJSON
         * @function swich Excel to JSON; 
         * return 
         */
        private string ExcelToJSON()
        {
            string text = "";
            int totalRows = ThisAddIn.Instance.workBook.ActiveSheet.UsedRange.Rows.Count;
            int totalColumns = ThisAddIn.Instance.workBook.ActiveSheet.UsedRange.Columns.Count;
            text += "[{";
            int startRow = 1;
            int startColumn = 1;
            for (int i = startRow+1; i != totalRows + 1; ++i)
            {
                for (int j = startColumn; j != totalColumns + 1; ++j)
                {               
                        text += "\"" + ThisAddIn.Instance.workBook.ActiveSheet.Cells(startRow, j).Value + "\":";
                        // In the last row and last column, modify text style;
                        if (i == totalRows && j == totalColumns)
                        {
                            text += "\"" + ThisAddIn.Instance.workBook.ActiveSheet.Cells(i, j).Value + "\"";
                        }
                        else
                        {
                            // Every Row Start, write "},";
                            if (j == totalColumns)
                            {
                                text += "\"" + ThisAddIn.Instance.workBook.ActiveSheet.Cells(i, j).Value + "\"},{";
                            }
                            else
                            {
                                text += "\"" + ThisAddIn.Instance.workBook.ActiveSheet.Cells(i, j).Value + "\",";
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
            Message(totalRows, totalColumns);
            return text;
        }

        /*
         * Write to file
         * @param path
         * @param name
         * @param text
         * return
         */
        public void WriteToFile(string path, string name, string text)
        {
            lock (this)
            {
                // File stream information;
                StreamWriter sw;
                FileInfo t = new FileInfo(path+name);
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
        private void Message(int totalRows, int totalColumns)
        {
            MessageBox.Show(
                "\n"
                +"  Excel to JSON\n"
                +"  Created by OneLei.\n"
                +"  Email: ahleiwolong@163.com \n"
                +"  Copyright (c) 2014 Year. All rights reserved.\n\n"
                +"  Save Successful!\n\n"
                +"  Total: " + totalRows + " Rows"
                +", " + totalColumns + " Columns");
        }

    }
}
