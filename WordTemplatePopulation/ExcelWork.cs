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
            data = new string[m - 1, n];//m-1 потому что m,n - размер всей таблицы, включая заголовки столбцов
        }
    }
    class ExcelWork
    {
        /// <summary>
        /// 
        /// </summary>
        /// <returns>filepath of chosen file</returns>
        public static string ChooseFile()
        {
            Stream myStream = null;
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.InitialDirectory = @"E:\GitHub\WordTemplatePopulation";
            //openFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            //openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;
            string file = null;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                file = openFileDialog1.FileName;
            }
            return file;
        }
        /// <summary>
        /// Opens, read Excel table and save data to ExcelData
        /// </summary>
        /// <returns>data from Excel table or null if table is empty or was error during opening</returns>
        public static ExcelData ReadExcelFile(string filePath)
        {

            if (filePath == null)
                return null;

            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            var ExcelApp = new Excel.Application();
            CultureInfo oldCi = System.Threading.Thread.CurrentThread.CurrentCulture;

            //!!!!!!!!!!!!!!!Делать проверку на правильность открытия Open
            Excel.Workbook wrkBook = ExcelApp.Workbooks.Open(filePath);
            if (wrkBook == null)
            {
                MessageBox.Show("Не удалось открыть WorkBook");
                return null;
            }


            Excel.Worksheet wrkSheet = wrkBook.Worksheets[1];
            //System.Threading.Thread.CurrentThread.CurrentCulture = oldCi;
            Excel.Range usedRange = wrkSheet.UsedRange;
            //Тут делать ExcelData(n), т.е. найти n из usedRange
            ExcelData excelTable = new ExcelData(usedRange.Rows.Count, usedRange.Columns.Count);

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
                return null;
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
            for (int i = 0; i < dataRange.Rows.Count; i++)
            {
                for (int j = 0; j < dataRange.Columns.Count; j++)
                {
                    excelTable.data[i, j] = Convert.ToString(dataRange.Cells[i + 1, j + 1].Value2);
                }
            }
            ExcelApp.Quit();
            return excelTable;
        }
    }
    class WordWork
    {
        //экземпляр ворда

        public static string WordPath;
        public static string ResultPath;
        public static void CreateDir(string path)
        {
            
            int a = path.LastIndexOf("\\");
            string name = path.Remove(0, a+1);
            string dir = path.Remove(a)+"\\Result_"+name;
            Directory.CreateDirectory(dir);
            ResultPath=dir;
    }

    public static string openWordFile()
        {
            //Stream myStream = null;
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.InitialDirectory = @"E:\GitHub\WordTemplatePopulation";
            
            //TODO: фильтр только doc, docx файлов
            openFileDialog1.Filter = "Word files|*.doc;*.docx";
            //openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;
            string file = null;
            
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                file = openFileDialog1.FileName;
            }
            CreateDir(file);
            WordPath = file;
            return file;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="maskSymbol">Symbol enclosing inserting field</param>
        /// <param name="excelData">data from Excel table</param>
        /// <param name="wordFilePath"></param>
        public static void CreateOutputFiles(string maskSymbol, ExcelData excelData, string wordFilePath)
        {
            //1.Open Word File for read

            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            //var WordApp = new Word.Application();

            var WordApp = new Microsoft.Office.Interop.Word.Application();
            //WordApp.Visible = true;

            // TODO: перенести это открытие файла после создания внутрь двух for
            Word.Document WordDoc = WordApp.Documents.Open(wordFilePath);
            Word.Range rng = WordDoc.Content;
            rng.Find.ClearFormatting();
            rng.Find.Forward = true;
            rng.Find.Replacement.ClearFormatting();
            //До TODO выше всё перенести внутрь for

            for (int i=0;i<excelData.data.GetLength(0);i++)
            {
                for (int j = 0; j < excelData.data.GetLength(1); j++)
                {
                    // TODO: Делать создание копии изначального файла. Имя файла = Полю "Имя или ФИО". Брать из excelData.ColumnTitles
                    //Add mask symbols to searched field
                    rng.Find.Text = maskSymbol + excelData.ColumnTitles[j] + maskSymbol;
                    rng.Find.Replacement.Text = excelData.data[i, j];
                    rng.Find.MatchCase = false;
                    rng.Find.Execute();

                    // TODO: Сохранить изменённый файл
                    //
                    
                }
            }

            


            //WordDoc = WordApp.Documents.Open(wordFilePath);
            //For excelData
            //Create Copy in our new directory
            //Fill @text@
            //save
            //End For
            WordApp.Quit();
        }




        //ExcelData excelTable = ExcelWork.OpenExcelFile();



        // System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
        //        var ExcelApp = new Excel.Application();
    }
}
    
    
      

