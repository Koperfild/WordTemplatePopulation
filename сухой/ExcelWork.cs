using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using System.Windows.Forms;

namespace WordTemplatePopulation
{
    public class ExcelData
    {
        public string[] ColumnTitles;
        public string[,] data;
        public ExcelData() { }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="n">Size of table</param>
        public ExcelData(int m, int n)
        {
            ColumnTitles = new string[n];
            data = new string[m,n];
        }
    }
    class ExcelWork
    {
        public static void OpenExcelFile()
        {
            Stream myStream = null;
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.InitialDirectory = @"C:\Users\Sergey\Desktop\женёе\WordTemplatePopulation";
            //openFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            //openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string file = openFileDialog1.FileName;
                var ExcelApp = new Excel.Application();
                Excel.Workbook wrkBook = ExcelApp.Workbooks.Open(file);
                Excel.Worksheet wrkSheet = wrkBook.Worksheets[1];
                Excel.Range usedRange = wrkSheet.UsedRange;
                ExcelData excelTable = new ExcelData();
                //========================СЧИТЫВАНИЕ ДАННЫХ ОТСЮДА====================
                List<List<string>> rawValueList = new List<List<string>>();
                int i = 0;
                foreach (Excel.Range row in usedRange.Rows)
                {
                    rawValueList.Add(new List<string>());
   
                    foreach(dynamic a in row.Value2)
                    {
                        if(a==null)
                        {
                            rawValueList[i].Add("");
                        }
                        else
                        rawValueList[i].Add(a.ToString());
                    }
                    i++;
                }
                wrkBook.Close();
                //========================СЧИТЫВАНИЕ ДАННЫХ ДОСЮДА====================
               
            }
            
        }
        

    }
}
