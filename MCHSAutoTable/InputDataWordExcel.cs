using MCHSAutoTable.Entityes.edds;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace MCHSAutoTable
{
    public class InputDataWordExcel
    {

        public void createTableWordEDDS(string FIO, List<string[]> EDDSTableList)
        {
            //Получение даты и времени
            DateTime dateTime = DateTime.Now;
            string time = dateTime.ToString("HH-mm");
            string date = dateTime.ToShortDateString();

            //Создание директории
            var thTH = new System.Globalization.CultureInfo("th-TH");
            string path = dateTime.Year.ToString()+"\\"+dateTime.Month.ToString()+"\\"+date;
            Directory.CreateDirectory(path);

            DocX doc = DocX.Create(path+"\\"+"Минский район_ЕДДС " + date + ".docx");

            //Шапка таблицы
            doc.InsertParagraph("Проведение проверок служб ЕДДС")
                .Font("Times New Roman")
                .FontSize(13)
                .Alignment = Xceed.Document.NET.Alignment.center;
            doc.InsertParagraph("СПИСОК СЛУЖБ Минского района (диспетчер " + FIO + ", " + date + ")")
                .Font("Times New Roman")
                .FontSize(13)
                .Alignment = Xceed.Document.NET.Alignment.center;

            if (EDDSTableList.Count != 0)
            {
                //Создаем таблицу
                Table table = doc.AddTable(EDDSTableList.Count+1, 6);
                table.Alignment = Alignment.center;

                // Шапка таблицы
                table.Rows[0].Cells[0].Paragraphs[0].Append("№ \nп/п").Font("Times New Roman").Alignment = Alignment.center;
                table.Rows[0].Cells[1].Paragraphs[0].Append("Время \nпроверки").Font("Times New Roman").Alignment = Alignment.center;
                table.Rows[0].Cells[2].Paragraphs[0].Append("Наименование \nорганизации\n(службы)").Font("Times New Roman").Alignment = Alignment.center;
                table.Rows[0].Cells[3].Paragraphs[0].Append("Техническое \nсостояние \nисправна/\nне исправна").Font("Times New Roman").Alignment = Alignment.center;
                table.Rows[0].Cells[4].Paragraphs[0].Append("Ф.И.О. \nпринявшего \nзвонок").Font("Times New Roman").Alignment = Alignment.center;
                table.Rows[0].Cells[5].Paragraphs[0].Append("Примечание").Font("Times New Roman").Alignment = Alignment.center;

                //Заполнение ячеек таблицы
                
                int i = 1;
                foreach (string[] edds in EDDSTableList)
                {
                    table.Rows[i].Cells[0].Paragraphs[0].Append((i).ToString()).Font("Times New Roman").Alignment = Alignment.center;
                    table.Rows[i].Cells[1].Paragraphs[0].Append(edds[0]).Font("Times New Roman").Alignment = Alignment.center;
                    table.Rows[i].Cells[2].Paragraphs[0].Append(edds[1]).Font("Times New Roman").Alignment = Alignment.center;
                    table.Rows[i].Cells[3].Paragraphs[0].Append(edds[2]).Font("Times New Roman").Alignment = Alignment.center;
                    table.Rows[i].Cells[4].Paragraphs[0].Append(edds[3]).Font("Times New Roman").Alignment = Alignment.center;
                    table.Rows[i].Cells[5].Paragraphs[0].Append(edds[4]).Font("Times New Roman").Alignment = Alignment.center;
                    i++;
                }


                table.AutoFit = AutoFit.Contents;
                doc.InsertParagraph().InsertTableAfterSelf(table);
                doc.Save();
            }
            else
                MessageBox.Show("Таблица не содержит данных, заполните данные во вкладке ЗАПОЛНЕНИЕ ЕДДС");
        }
    }
}
