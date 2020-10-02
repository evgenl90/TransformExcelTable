using System;
using System.IO;
using ExcelPro = Microsoft.Office.Interop.Excel;

namespace Excel
{
    class Program
    {

        static void Main(string[] args)
        {

    
            /* создание файла excel*/
            ExcelPro.Application ObjWorkExcel = new ExcelPro.Application();
            ExcelPro.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Add(); // создание книги
            ExcelPro.Worksheet worksheet = ObjWorkBook.Worksheets[1];                                                      
           
            worksheet.Cells[1, "A"] = "Обозначение";
            worksheet.Cells[1, "B"] = "Наименование";
            worksheet.Cells[1, "C"] = "Количество";
            worksheet.Cells[1, "D"] = "Исполнение";
            worksheet.Cells[1, "E"] = "Примечание";

            worksheet.Columns[1].ColumnWidth = 25;
            worksheet.Columns[2].ColumnWidth = 30;
            worksheet.Columns[3].ColumnWidth = 15;
            worksheet.Columns[4].ColumnWidth = 25;
            worksheet.Columns[5].ColumnWidth = 25;

            worksheet.Columns.RowHeight = 20;

            ExcelPro.Range rng = (ExcelPro.Range)worksheet.Rows[1];
            rng.Font.Bold = true;

            /*получение имен файлов*/
            Console.WriteLine("Получение имен файлов...");
            string dirName = @"C:\excel_convert\";

            if (Directory.Exists(dirName))
            {
                string[] files = Directory.GetFiles(dirName);
                foreach (string item_file in files)
                {


                    /*Считывание документа*/
                    ExcelPro.Range Rng;
                    ExcelPro.Workbook xlWB;
                    ExcelPro.Worksheet xlSht;
                    int iLastRow, iLastCol;

                    ExcelPro.Application xlApp = new ExcelPro.Application(); //создаём приложение Excel
                    xlWB = xlApp.Workbooks.Open(item_file,
                                 Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                 false, Type.Missing, Type.Missing, Type.Missing,
                                 Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                 Type.Missing, Type.Missing); //открываем наш файл           
                    xlSht = xlWB.Worksheets[1]; //или так xlSht = xlWB.ActiveSheet //активный лист

                    iLastRow = xlSht.Cells[xlSht.Rows.Count, "BD"].End[ExcelPro.XlDirection.xlUp].Row; //последняя заполненная строка в столбце А
                    iLastCol = xlSht.Cells[iLastRow, xlSht.Columns.Count].End[ExcelPro.XlDirection.xlToLeft].Column; //последний заполненный столбец в 1-й строке

                    Rng = (ExcelPro.Range)xlSht.Range["L1", xlSht.Cells[iLastRow, iLastCol]]; // записи диапазона ячеек в переменную Rng


                    var dataArr = (object[,])Rng.Value; //чтение данных из ячеек в массив\
                    int rows = dataArr.GetUpperBound(0) + 1;
                    int columns = dataArr.Length / rows;

                    bool flag1 = false;
                    bool flag2 = false;
                    int countRow = worksheet.Cells[worksheet.Rows.Count, "B"].End[ExcelPro.XlDirection.xlUp].Row + 2; //последняя заполненная строка в столбце
                    
                    for (int i = 1; i < rows; i++)
                    {

                        if (Convert.ToString(dataArr[i, 24]) == "" || dataArr[i, 24] == null) continue;

                            if (Convert.ToString(dataArr[i, 24]) == "Детали")
                        {
                            flag1 = true;
                            continue;
                        }

                        if (Convert.ToString(dataArr[i, 24]) == "Стандартные изделия")
                        {
                            flag1 = false;
                            flag2 = true;
                            continue;
                        }

                        if (flag1 || flag2)
                        {
                            worksheet.Cells[countRow, "A"] = Convert.ToString(dataArr[i, 1]);
                            worksheet.Cells[countRow, "B"] = Convert.ToString(dataArr[i, 24]);
                            worksheet.Cells[countRow, "C"] = Convert.ToString(dataArr[i, 45]);

                            countRow++;

                        }

                       
                    }

                    //закрытие Excel
                    xlWB.Close(true); //сохраняем и закрываем файл
                    xlApp.Quit();

                    Console.WriteLine("Файл " + item_file + " обработан...");

                }
            }

            //ObjWorkExcel.Visible = true;
            ObjWorkExcel.Application.ActiveWorkbook.SaveAs(@"C:\excel_convert\success\success.xlsx", Type.Missing,
   Type.Missing, Type.Missing, Type.Missing, false, ExcelPro.XlSaveAsAccessMode.xlNoChange,
   Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            Console.WriteLine("Файл success успешно записан!");
            Console.WriteLine("Для завершения нажмите клавишу Enter...");
            Console.ReadKey();


        }
    }


}
