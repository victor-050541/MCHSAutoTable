using MCHSAutoTable.Entityes.edds;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace MCHSAutoTable
{
    public class InputDataWordExcel
    {
        /* public void printInWordEDDS(string nameOperator)
         {
             using (ApplicationContextTableEDDS db = new ApplicationContextTableEDDS())
             {
                 DateTime dateTime = DateTime.Now;
                 string date = dateTime.ToShortDateString();

                 // получаем объекты из бд и выводим на консоль
                 List<TableEDDS> listEDDS= db.TableEDDSlist.ToList();
                 //создание документа
                 Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
                 //загрузка документа
                 Microsoft.Office.Interop.Word.Document doc = app.Documents.Open("testdoc1.docx");
                 object missing = System.Reflection.Missing.Value;

                 Range r = doc.Range();
                 doc.Content.Text += "Проведение проверок служб ЕДДС\r\nСПИСОК СЛУЖБ Минского района (диспетчер " +
                     nameOperator + " " + date + ")\r\n";
                 app.Visible = true;


                 Table t = doc.Tables.Add(r, listEDDS.Count, 4);
                 t.Borders.Enable = 1;

                 for (int i = 1; i <= listEDDS.Count; i++)
                 {
                     t.Cell(i, 1).Range.Text = i.ToString()+".";
                     t.Cell(i, 2).Range.Text = listEDDS[i - 1].Time;
                     // t.Cell(i, 3).Range.Text = listEDDS[i - 1].listEDDS[i-1].Name;                    
                     t.Cell(i, 3).Range.Text = listEDDS[i - 1].Working;
                     t.Cell(i, 4).Range.Text = listEDDS[i - 1].FIO;
                     //t.Cell(i, 6).Range.Text = listEDDS[i - 1].listEDDS[i - 1].PhoneNumber;                    
                 }

                 doc.Save();
                 //app.Documents.Open(@"C:/String/Colom.doc");
                 doc.Close();
                 app.Quit();
             }                       
         }
     }*/

        /*public void HelloWorld(string documentFileName)
        {
            // Create a Wordprocessing document. 
            using (WordprocessingDocument myDoc =
                   WordprocessingDocument.Create(documentFileName,
                                 WordprocessingDocumentType.Document))
            {
                // Add a new main document part. 
                MainDocumentPart mainPart = myDoc.AddMainDocumentPart();
                //Create Document tree for simple document. 
                mainPart.Document = new Document();
                //Create Body (this element contains
                //other elements that we want to include 
                Body body = new Body();
                //Create paragraph 
                Paragraph paragraph = new Paragraph();
                Run run_paragraph = new Run();
                // we want to put that text into the output document 
                Text text_paragraph = new Text("Hello World!");
                //Append elements appropriately. 
                run_paragraph.Append(text_paragraph);
                paragraph.Append(run_paragraph);
                body.Append(paragraph);
                mainPart.Document.Append(body);
                // Save changes to the main document part. 
                mainPart.Document.Save();
            }
        }*/

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

    /*public void createTableWordPatient(List<string[]> PatientTableList)
    {
        DateTime dateTime = DateTime.Now;
        string date = dateTime.ToShortDateString();

        DocX doc = DocX.Create(date + " Минская обл сведения по больничным новая форма.docx");

        //Верхний текст
        doc.InsertParagraph("Список работников\n" + "Минского областного управления МЧС,\n"
            + "находящихся на справках о временной нетрудоспособности\n"
            + "по состоянию на " + date + " года")
            .Font("Times New Roman")
            .FontSize(13)
            .Alignment = Xceed.Document.NET.Alignment.center;

        doc.PageWidth = 1000;


        int countString = 0;

        if (countString != 0)
        {
            //Создаем таблицу
            Table table = doc.AddTable(countString + 1, 10);
            table.Alignment = Alignment.center;
            table.Alignment = Alignment.center;
            table.AutoFit = AutoFit.Window;
            // заполнение ячейки текстом
            table.Rows[0].MergeCells(7, 8);
            table.Rows[0].Cells[7].Paragraphs[0].Append("Лечение").Font("Times New Roman").Bold().FontSize(13).Alignment = Alignment.center;

            table.Rows[0].Cells[0].Paragraphs[0].Append("№ п/п").Font("Times New Roman").Bold().FontSize(13).Alignment = Alignment.center;
            table.Rows[0].Cells[1].Paragraphs[0].Append("ФИО\nработника").Font("Times New Roman").Bold().FontSize(13).Alignment = Alignment.center;
            table.Rows[0].Cells[2].Paragraphs[0].Append("Подразде\nление").Font("Times New Roman").Bold().FontSize(13).Alignment = Alignment.center;
            table.Rows[0].Cells[3].Paragraphs[0].Append("Должность, звание").Font("Times New Roman").Bold().FontSize(13).Alignment = Alignment.center;
            table.Rows[0].Cells[4].Paragraphs[0].Append("Начало\nвременной\nнетрудоспособности").Font("Times New Roman").Bold().FontSize(13).Alignment = Alignment.center;
            table.Rows[0].Cells[5].Paragraphs[0].Append("Предварительный/\nокончательный\nдиагноз").Font("Times New Roman").Bold().FontSize(13).Alignment = Alignment.center;
            table.Rows[0].Cells[6].Paragraphs[0].Append("Номер\nдежурной\nсмены").Font("Times New Roman").Bold().FontSize(13).Alignment = Alignment.center;

            table.Rows[1].Cells[7].Paragraphs[0].Append("Амб\nулат\nорно\nе").Font("Times New Roman").Bold().FontSize(13).Alignment = Alignment.center;
            table.Rows[1].Cells[8].Paragraphs[0].Append("Стац\nиона\nрное").Font("Times New Roman").Bold().FontSize(13).Alignment = Alignment.center;
            table.Rows[0].Cells[8].Paragraphs[0].Append("Кто и\nкогда\nпосещал").Font("Times New Roman").Bold().FontSize(13).Alignment = Alignment.center;

            table.Rows[2].MergeCells(0, 10);
            table.Rows[2].Cells[0].Paragraphs[0].Append("Минское областное управление МЧС").Font("Times New Roman").Bold().FontSize(13).Alignment = Alignment.center;
            table.Rows[3].MergeCells(0, 10);
            table.Rows[3].Cells[0].Paragraphs[0].Append("Минский РОЧС").Font("Times New Roman").Bold().UnderlineColor(Color.Yellow).
                FontSize(13).Alignment = Alignment.center;

            int count = 1;
            int i = 4;


            //table.Rows[i].Cells[0].Paragraphs[0].Append(count + ".").Font("Times New Roman").FontSize(13).Alignment = Alignment.center;
            table.Rows[i].Cells[1].Paragraphs[0].Append(Convert.ToString(sqlDataReader2["patient_name"])).Font("Times New Roman").FontSize(13).Alignment = Alignment.center;
            table.Rows[i].Cells[2].Paragraphs[0].Append(Convert.ToString(sqlDataReader2["position_patient"]) + ", " + Convert.ToString(sqlDataReader["shift_patient"])).Font("Times New Roman").FontSize(13).Alignment = Alignment.center;
            table.Rows[i].Cells[3].Paragraphs[0].Append(Convert.ToString(sqlDataReader2["division"])).Font("Times New Roman").FontSize(13).Alignment = Alignment.center;
            table.Rows[i].Cells[4].Paragraphs[0].Append(Convert.ToString(sqlDataReader2["date"])).Font("Times New Roman").FontSize(13).Alignment = Alignment.center;
            table.Rows[i].Cells[5].Paragraphs[0].Append(Convert.ToString(sqlDataReader2["diagnosis"])).Font("Times New Roman").FontSize(13).Alignment = Alignment.center;
            table.Rows[i].Cells[6].Paragraphs[0].Append("1").Font("Times New Roman").FontSize(13).Alignment = Alignment.center;
            if (Convert.ToString(sqlDataReader2["healing"]).Equals("Амбулаторное"))
            {
                table.Rows[i].Cells[7].Paragraphs[0].Append("+").Font("Times New Roman").FontSize(13).Alignment = Alignment.center;
                table.Rows[i].Cells[8].Paragraphs[0].Append("-").Font("Times New Roman").FontSize(13).Alignment = Alignment.center;
            }
            else
            {
                table.Rows[i].Cells[7].Paragraphs[0].Append("-").Font("Times New Roman").FontSize(13).Alignment = Alignment.center;
                table.Rows[i].Cells[8].Paragraphs[0].Append("+").Font("Times New Roman").FontSize(13).Alignment = Alignment.center;
            }
            table.Rows[i].Cells[9].Paragraphs[0].Append(Convert.ToString(sqlDataReader2["who"])).Font("Times New Roman").FontSize(13).Alignment = Alignment.center;


            table.AutoFit = AutoFit.Contents;
            doc.InsertParagraph().InsertTableAfterSelf(table);
            doc.Save();
        }
        else
            MessageBox.Show("Таблица не содержит данных, заполните данные во вкладке ЗАПОЛНЕНИЕ ЕДДС");
    }          */  
}
