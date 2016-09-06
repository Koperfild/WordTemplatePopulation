using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using System.Windows.Forms;
using System.Globalization;

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
            data = new string[m-1,n];//m-1 потому что m,n - размер всей таблицы, включая заголовки столбцов
        }
    }
    class ExcelWork
    {
        public static void OpenExcelFile()
        {
            Stream myStream = null;
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.InitialDirectory = @"E:\GitHub\WordTemplatePopulation";
            //openFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            //openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string file = openFileDialog1.FileName;

                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                var ExcelApp = new Excel.Application();
                CultureInfo oldCi = System.Threading.Thread.CurrentThread.CurrentCulture;

                //!!!!!!!!!!!!!!!Делать проверку на правильность открытия Open
                Excel.Workbook wrkBook = ExcelApp.Workbooks.Open(file);
                

                Excel.Worksheet wrkSheet = wrkBook.Worksheets[1];
                //System.Threading.Thread.CurrentThread.CurrentCulture = oldCi;
                Excel.Range usedRange = wrkSheet.UsedRange;
                //Тут делать ExcelData(n), т.е. найти n из usedRange
                ExcelData excelTable = new ExcelData(usedRange.Rows.Count,usedRange.Columns.Count);

                #region ~~~~~~~~~~Experiments~~~~~~~~~~~
                //~~~~~~~~~~~~~~~~~~~~~~~~~~~Experiments~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

                //int m = usedRange.Rows.Count;
                //int n = usedRange.Columns.Count;
                //Excel.Range tmp1 = usedRange.Offset[1, 0];
                //string tmp_1 = usedRange.Cells[2, 3].Value2.ToString();
                //string tmp_2 = tmp1.Cells[2, 3].Value2.ToString();
                //m = usedRange.Rows.Count;
                //n = usedRange.Columns.Count;
                //tmp_1 = Convert.ToString(usedRange.Cells[usedRange.Rows.Count, usedRange.Columns.Count].Value2);
                //tmp_2 = Convert.ToString(tmp1.Cells[usedRange.Rows.Count-1, usedRange.Columns.Count-1].Value2);
                //tmp_2 = Convert.ToString(tmp1.Cells[usedRange.Rows.Count, usedRange.Columns.Count].Value2);
                //tmp1 = usedRange.Resize[usedRange.Rows.Count - 1, Type.Missing];

                //~~~~~~~~~~~~~~~~~~~~~~~~~~~Experiments~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                #endregion

                //Если таблица пустая
                if (usedRange.Rows.Count <= 1)
                {
                    MessageBox.Show("Вы ввели пустую таблицу Excel, выберите другой файл Excel");
                    wrkBook.Close();//Проверять метод Close. Т.к. могут быть изменения в файле при простом открытии файла
                    if (ExcelApp != null) ExcelApp.Quit();
                }

                for (int i = 0; i < usedRange.Columns.Count; i++)
                {
                    excelTable.ColumnTitles[i] = Convert.ToString(usedRange.Cells[1, i + 1].Value2);
                }

                //Сдвигаем диапазон вниз на 1 строку
                Excel.Range dataRange = usedRange.Offset[1, 0];
                //Убираем получившийся лишним при сдвиге вниз последний ряд
                dataRange = dataRange.Resize[usedRange.Rows.Count - 1, Type.Missing];

                //Заполняем таблицу данных данными
                for (int i=0;i<dataRange.Rows.Count;i++)
                {
                    for (int j=0;j<dataRange.Columns.Count;j++)
                    {
                        excelTable.data[i,j] = Convert.ToString(dataRange.Cells[i+1,j+1].Value2);
                    }                    
                }               
            }
            
        }      
    }
}
