using System.Runtime.InteropServices;
using Xceed.Document.NET;
using Xceed.Words.NET;
using Excel = Microsoft.Office.Interop.Excel;


namespace MCHSAutoTable
{
    public class InputDataWordExcel
    {
        public static void CreateTableWordEdds(string fio, List<string[]> eddsTableList)
        {
            //Получение даты и времени
            DateTime dateTime = DateTime.Now;
            string time = dateTime.ToString("HH-mm");
            string date = dateTime.ToShortDateString();

            //Создание директории
            var thTh = new System.Globalization.CultureInfo("th-TH");
            string path = dateTime.Year.ToString() + "\\" + dateTime.Month.ToString() + "\\" + date;

            Directory.CreateDirectory(path);

            DocX doc = DocX.Create(path + "\\" + "Минский район_ЕДДС " + date + ".docx");

            //Шапка таблицы
            doc.InsertParagraph("Проведение проверок служб ЕДДС")
                .Font("Times New Roman")
                .FontSize(13)
                .Alignment = Xceed.Document.NET.Alignment.center;
            doc.InsertParagraph("СПИСОК СЛУЖБ Минского района (диспетчер " + fio + ", " + date + ")")
                .Font("Times New Roman")
                .FontSize(13)
                .Alignment = Xceed.Document.NET.Alignment.center;

            if (eddsTableList.Count != 0)
            {
                //Создаем таблицу
                Table table = doc.AddTable(eddsTableList.Count + 1, 6);
                table.Alignment = Alignment.center;

                // Шапка таблицы
                table.Rows[0].Cells[0].Paragraphs[0].Append("№ \nп/п").Font("Times New Roman").Alignment =
                    Alignment.center;
                table.Rows[0].Cells[1].Paragraphs[0].Append("Время \nпроверки").Font("Times New Roman").Alignment =
                    Alignment.center;
                table.Rows[0].Cells[2].Paragraphs[0].Append("Наименование \nорганизации\n(службы)")
                    .Font("Times New Roman").Alignment = Alignment.center;
                table.Rows[0].Cells[3].Paragraphs[0].Append("Техническое \nсостояние \nисправна/\nне исправна")
                    .Font("Times New Roman").Alignment = Alignment.center;
                table.Rows[0].Cells[4].Paragraphs[0].Append("Ф.И.О. \nпринявшего \nзвонок").Font("Times New Roman")
                    .Alignment = Alignment.center;
                table.Rows[0].Cells[5].Paragraphs[0].Append("Примечание").Font("Times New Roman").Alignment =
                    Alignment.center;

                //Заполнение ячеек таблицы

                int i = 1;
                foreach (string[] edds in eddsTableList)
                {
                    table.Rows[i].Cells[0].Paragraphs[0].Append((i).ToString()).Font("Times New Roman").Alignment =
                        Alignment.center;
                    table.Rows[i].Cells[1].Paragraphs[0].Append(edds[0]).Font("Times New Roman").Alignment =
                        Alignment.center;
                    table.Rows[i].Cells[2].Paragraphs[0].Append(edds[1]).Font("Times New Roman").Alignment =
                        Alignment.center;
                    table.Rows[i].Cells[3].Paragraphs[0].Append(edds[2]).Font("Times New Roman").Alignment =
                        Alignment.center;
                    table.Rows[i].Cells[4].Paragraphs[0].Append(edds[3]).Font("Times New Roman").Alignment =
                        Alignment.center;
                    table.Rows[i].Cells[5].Paragraphs[0].Append(edds[4]).Font("Times New Roman").Alignment =
                        Alignment.center;
                    i++;
                }


                table.AutoFit = AutoFit.Contents;
                doc.InsertParagraph().InsertTableAfterSelf(table);
                doc.Save();
            }
            else
                MessageBox.Show("Таблица не содержит данных, заполните данные во вкладке ЗАПОЛНЕНИЕ ЕДДС");
        }

        public void CreateTableExcelPatients()
        {
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
                return;
            }

            //Получение даты и времени
            DateTime dateTime = DateTime.Now;
            string time = dateTime.ToString("HH-mm");
            string date = dateTime.ToShortDateString();
            //Создание директории
            var thTh = new System.Globalization.CultureInfo("th-TH");
            string path = dateTime.Year.ToString() + "\\" + dateTime.Month.ToString() + "\\" + date;
            Directory.CreateDirectory(path);

            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            xlWorkSheet.Cells[1, 1] = "Сведения о работниках, находящихся на больничном 08.30 ежедневно";
            xlWorkSheet.Cells[2, 1] = "Наименование подразделений";
            xlWorkSheet.Cells[2, 2] = "по штату";
            xlWorkSheet.Cells[2, 3] = "больные";
            xlWorkSheet.Cells[2, 4] = "сменный график работы";
            xlWorkSheet.Cells[2, 5] = "ежедневный график работы";
            xlWorkSheet.Cells[2, 6] = "гражданские работники";
            xlWorkSheet.Cells[2, 7] = "из них";
            xlWorkSheet.Cells[2, 8] = "руководящий состав подразделений";

            var fullPath = path + "\\" + "Минский район сведения по больным " + date + ".xls";
            xlWorkBook.SaveAs(
                fullPath,
                Excel.XlFileFormat.xlWorkbookNormal,
                misValue,
                misValue,
                misValue,
                misValue,
                Excel.XlSaveAsAccessMode.xlExclusive,
                misValue,
                misValue,
                misValue,
                misValue,
                misValue
            );
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
        }
    }
}