using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Office.CustomUI;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Vml;
using MCHSAutoTable.Entityes;
using MCHSAutoTable.Entityes.coworker;
using MCHSAutoTable.Entityes.edds;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.ApplicationServices;
using System.Reflection.Emit;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TaskbarClock;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ToolBar;
using Button = System.Windows.Forms.Button;
using Color = System.Drawing.Color;

namespace MCHSAutoTable
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            dataEDDS = workingDB.getEDDS();            
        }

        //Статус кнопок для изменения цвета и закрытия панелей
        byte btn1Status = 0;
        byte btn2Status = 0;
        byte btn3Status = 0;
        byte btn4Status = 0;

        int indexDataEDDS = 0;

        //Для работы с БД
        WorkingDB workingDB = new WorkingDB();

        //Вывод информации из БД в WORD или Excel
        InputDataWordExcel inputData = new InputDataWordExcel();

        //ЕДДС
        private void button1_Click(object sender, EventArgs e)
        {
            //inputData.createTableExcelPatients();
            workingDB.clearRowInTableEDDS();
            listBox1.Items.Clear();

            //InputDataWordExcel inputDataWordExcel = new InputDataWordExcel();
            //inputDataWordExcel.createTableExcelPatients();

            //Загрузка в список
            if (dataEDDS.Count > 0)
            {
                dataEDDS = workingDB.getEDDS();
                label3.Text = "Номер телефона: " + dataEDDS[indexDataEDDS][1];
                foreach (string[] s in dataEDDS)
                {
                    listBox1.Items.Add(s[0]);
                }
                radioButton1.Checked = true;
                listBox1.SetSelected(0, true);

                btn1Status = buttonColorAndOpenPanel(button1, btn1Status, panel1);
            }
            else
                MessageBox.Show("Заполните данные в разделе Ведомства ЕДДС");
        }

        //Времено нетрудоспособные
        private void button2_Click(object sender, EventArgs e)
        {
            listBox2.Items.Clear();
            updateDataStaff();
            updateDataDiagnosis();
            updateTablePatient();
            btn2Status = buttonColorAndOpenPanel(button2, btn2Status, panel6);
        }        
        
        //Ведомства ЕДДС
        private void button4_Click(object sender, EventArgs e)
        {
            btn3Status = buttonColorAndOpenPanel(button4, btn3Status, panel2);

            //Обновление БД
            updateTableEDDS();
        }

        //Работники
        private void button6_Click(object sender, EventArgs e)
        {
            comboBox2.Items.Clear();
            comboBox4.Items.Clear();
            comboBox5.Items.Clear();

            updateDataPosition();
            updateDataSubDepartment();
            updateDataStaff();

            btn4Status = buttonColorAndOpenPanel(button6, btn4Status, panel3);
            List<String> listRank = new List<String>() { "рядовой", "мл. сержант", "сержант", "ст. сержант", "старшина",
             "прапорщик", "ст. прапорщик", "мл. лейтенант", "лейтенант", "ст. лейтенант", "капитан", "майор", "подполковник", "полковник"};

            List<String> listShift = new List<String>() { "-", "1", "2", "3"};

            comboBox6.DataSource = listShift;
            comboBox3.DataSource = listRank;
        }        

        //Обновление таблицы ведомства ЕДДС
        List<string[]> dataEDDS = new List<string[]>();
        void updateTableEDDS()
        {            
            dataGridView1.Rows.Clear();
            dataEDDS.Clear();
            comboBox1.Items.Clear();
            comboBox1.ResetText();
            textBox4.Clear();
            maskedTextBox1.Clear();            

            dataEDDS = workingDB.getEDDS();

            if (dataEDDS.Count > 0)
            {                
                foreach (string[] s in dataEDDS)
                {
                    dataGridView1.Rows.Add(s);
                    comboBox1.Items.Add(s[0]);
                }
            }                                
        }

        //Обновление таблицы ведомства ЕДДС
        List<string[]> dataPatients = new List<string[]>();
        void updateTablePatient()
        {
            dataGridView3.Rows.Clear();
            comboBox10.Items.Clear();

            dataPatients = workingDB.getPatients();

            if (dataPatients.Count > 0)
            {
                foreach (string[] s in dataPatients)
                {
                    dataGridView3.Rows.Add(s);
                    comboBox10.Items.Add(s[0]);
                }
            }
        }

        //Обновление checkBox с должностями
        List<string[]> dataPosition = new List<string[]>();
        void updateDataPosition()
        {
            dataPosition = workingDB.getPosition();

            comboBox2.Items.Clear();
            comboBox7.Items.Clear();

            if (dataPosition.Count > 0)
            {
                foreach (string[] s in dataPosition)
                {
                    comboBox2.Items.Add(s[0]);
                    comboBox7.Items.Add(s[0]);
                }
            }
        }

        //Обновление checkBox с подразделениями
        List<string[]> dataSubDep = new List<string[]>();
        void updateDataSubDepartment()
        {
            dataSubDep = workingDB.getSubDepartment();

            comboBox4.Items.Clear();
            comboBox7.Items.Clear();

            if (dataSubDep.Count > 0)
            {
                foreach (string[] s in dataSubDep)
                {
                    comboBox4.Items.Add(s[0]);
                    comboBox7.Items.Add(s[0]);
                } 
            }
        }

        //Обновление checkBox с сотрудниками
        List<string[]> dataStaff = new List<string[]>();
        void updateDataStaff()
        {
            dataGridView2.Rows.Clear();
            comboBox5.Items.Clear();
            dataStaff = workingDB.getStaff();

            if (dataStaff.Count > 0)
            {
                foreach (string[] s in dataStaff)
                {
                    dataGridView2.Rows.Add(s);
                    comboBox5.Items.Add(s[0]);
                    listBox2.Items.Add(s[0]);
                }
            }
        }

        //Обновление checkBox с Болезнями
        List<string[]> dataDiagnosis = new List<string[]>();
        void updateDataDiagnosis()
        {
            dataDiagnosis = workingDB.getDiagnosis();

            comboBox8.Items.Clear();
            comboBox9.Items.Clear();

            if (dataDiagnosis.Count > 0)
            {
                foreach (string[] s in dataDiagnosis)
                {
                    comboBox8.Items.Add(s[0]);
                    comboBox9.Items.Add(s[0]);
                }
            }
        }

        //Добавить ведомство ЕДДС
        private void button5_Click_1(object sender, EventArgs e)
        {
            string name = textBox4.Text;
            string phoneNumber = maskedTextBox1.Text;
            workingDB.addDBDepartmentEDDS(name, phoneNumber);
            updateTableEDDS();
        }

        //Удалить ведомство ЕДДС
        private void button7_Click(object sender, EventArgs e)
        {
            if (!(comboBox1.SelectedIndex == -1))
            {                
                workingDB.deleteEDDS(dataEDDS[comboBox1.SelectedIndex][2]);
                updateTableEDDS();
            }
            else
            {
                MessageBox.Show("Не выбран элемент для удаления!");
            }            
        }

        List<string[]> tableDataEDDS = new List<string[]>();
        //Добавить в таблицу ЕДДС для вывода
        private void button3_Click(object sender, EventArgs e)
        {                           
                //string date = dateTime.ToShortDateString();//Дата для файла
                //string edds = dataEDDS[listBox1.SelectedIndex][0];
                //string phoneNumberTable = dataEDDS[listBox1.SelectedIndex][1];
                
                DateTime dateTime = DateTime.Now;
                string time = dateTime.ToString("HH-mm");
                textBox1.Text = time;

                string working = radioButton1.Checked ? "Исправна" : "Не исправна";
                string fedds = textBox2.Text;
               
                //listBox1.SelectedIndex += 1;

                if (!textBox2.Text.Equals(""))
                {                    
                    label3.Text = "Номер телефона: " + dataEDDS[indexDataEDDS][1];
                    workingDB.addDBTableEDDS(fedds, time, working);

                    textBox2.Clear();

                    if (dataEDDS.Count() - 1 != listBox1.SelectedIndex)
                    {
                        listBox1.SelectedIndex += 1;
                    }
                    else
                    {
                        btn1Status = 0;
                        button1.BackColor = Color.ForestGreen;

                        button3.Enabled = false;
                        textBox2.Enabled = false;
                        panel5.Visible = true;
                    }
                }
                else
                    MessageBox.Show("Поле с Фамилией не заполено!");
                radioButton1.Checked = true;          
        }

        //Кнопка завершить и сохранить данные в БД ЕДДС
        private void button13_Click(object sender, EventArgs e)
        {
            tableDataEDDS = workingDB.getTableEDDS();
            //Вывод в WORD
            inputData.createTableWordEDDS(textBox6.Text, tableDataEDDS);
            panel1.Hide();
            button3.Enabled = true;
            panel5.Visible = false;
            textBox2.Enabled = true;
            textBox6.Clear();
        }

        //Смена цвета для кнопок МЕНЮ
        byte buttonColorAndOpenPanel(Button button, byte btn, Panel panel)
        {
            if (btn == 0)
            {
                invisibleAllPanel();
                button.BackColor = Color.DarkSlateGray;
                panel.Visible = true;
            }
            else
            {
                button.BackColor = Color.ForestGreen;
                panel.Visible = false;
            }
            btn = (byte)(btn == 0 ? 1 : 0);
            return btn;
        }

        //Смена номеров телефонов при выборе ЕДДС в разделе ЕДДС
        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            DateTime dateTime = DateTime.Now;
            string time = dateTime.ToString("HH-mm");
            textBox1.Text = time;

            string phoneNumberTable = dataEDDS[listBox1.SelectedIndex][1];
            label3.Text = "Номер телефона: " + phoneNumberTable;
        }

        //Скрытие всех панелей и кнопок МЕНЮ
        void invisibleAllPanel()
        {
            panel1.Visible = false;
            panel2.Visible = false;
            panel3.Visible = false;
            panel6.Visible = false;

            button1.BackColor = Color.ForestGreen;
            button2.BackColor = Color.ForestGreen;
            button4.BackColor = Color.ForestGreen;
            button6.BackColor = Color.ForestGreen;

            btn1Status = 0;
            btn2Status = 0;
            btn3Status = 0;
            btn4Status = 0;
        }

        int indexDataDBWorkStaff = 0;
        //Смена информации в Панели в разделе Работники
        public void messageBoxAddData(int paramSwitch)
        {
            switch (paramSwitch)
            {
                case 0:
                    label15.Text = "Добавление должности";
                    indexDataDBWorkStaff = 1;
                    updateDataPosition();
                    panel4.Visible = true;
                    break;
                case 1:
                    label15.Text = "Добавление подразделения";
                    indexDataDBWorkStaff = 0;
                    updateDataSubDepartment();
                    panel4.Visible = true;
                    break;
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            messageBoxAddData(0);
        }

        private void button10_Click(object sender, EventArgs e)
        {
            messageBoxAddData(1);
        }

        private void button12_Click(object sender, EventArgs e)
        {
            if (indexDataDBWorkStaff == 1)
            {
                workingDB.addDBPosition(textBox5.Text);
                comboBox2.Items.Clear();
                updateDataPosition();
            }
            else
            {
                workingDB.addDBSubDepartment(textBox5.Text);
                comboBox4.Items.Clear();
                updateDataSubDepartment();
            }
            panel4.Visible = false;
            textBox5.Clear();
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            panel4.Visible = false;
            comboBox7.Items.Clear();
            textBox5.Clear();
        }

        //Добавление сотрудника
        private void button9_Click(object sender, EventArgs e)
        {
            if (!(textBox3.Text.Equals("") || comboBox2.Text.Equals("") || comboBox4.Text.Equals("") || maskedTextBox2.Text.Equals("") || comboBox3.Text.Equals("") || comboBox6.Text.Equals("")))
            {
                workingDB.addDBStaff(textBox3.Text, comboBox2.Text, comboBox4.Text, maskedTextBox2.Text, comboBox3.Text, comboBox6.Text);

                textBox3.Clear();
                comboBox2.Text = "";
                comboBox4.Text = "";
                maskedTextBox2.Text = "";
                comboBox3.Text = "";
                comboBox6.Text = "";

                updateDataStaff();
            }
            else
            {
                MessageBox.Show("Не заполнены данные!");
            }
            
        }

        //Удаление сотрудника
        private void button8_Click(object sender, EventArgs e)
        {
            if (!(comboBox5.SelectedIndex == -1))
            {
                workingDB.deleteStaff(dataStaff[comboBox5.SelectedIndex][6]);
                updateDataStaff();
            }
            else
            {
                MessageBox.Show("Не выбран элемент для удаления!");
            }
        }

        //Удаление Должности или Подразделения
        private void button14_Click(object sender, EventArgs e)
        {
            if (!(comboBox7.SelectedIndex == -1))
            {
                if (indexDataDBWorkStaff == 1)
                {
                    workingDB.deletePosition(dataPosition[comboBox7.SelectedIndex][1]);
                    updateDataPosition();
                }
                else
                {
                    workingDB.deleteSubDep(dataSubDep[comboBox7.SelectedIndex][1]);
                    updateDataSubDepartment();
                }

                panel4.Visible = false;
            }
            else
            {
                MessageBox.Show("Не выбран элемент для удаления!");
            }
            comboBox7.Items.Clear();
            textBox5.Clear();
        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {
            string search = textBox9.Text;
            int index = listBox2.FindString(search);
            listBox2.SelectedIndex = index; 
        }

        //Открытие панели для добавление Диагноза
        private void button15_Click(object sender, EventArgs e)
        {
            updateDataDiagnosis();
            panel7.Visible = true;
        }

        //Добавление Диагноза в БД
        private void button18_Click(object sender, EventArgs e)
        {
            workingDB.addDBDiagnosis(textBox7.Text);
            updateDataDiagnosis();

            textBox7.Clear();
            comboBox9.Items.Clear();
            panel7.Visible = false;
        }

        //Удаление диагноза из БД
        private void button17_Click(object sender, EventArgs e)
        {
            if (!(comboBox9.SelectedIndex == -1))
            {
                workingDB.deleteDiagnosis(dataDiagnosis[comboBox9.SelectedIndex][1]);
                updateDataDiagnosis();
                panel7.Visible = false;
            }
            else
            {
                MessageBox.Show("Не выбран элемент для удаления!");
            }
            comboBox9.Items.Clear();
            textBox7.Clear();
        }

        //Закрытие панели для добавление диагноза
        private void pictureBox2_Click(object sender, EventArgs e)
        {
            panel7.Visible = false;
        }

        //Добавление Пациента
        private void button16_Click(object sender, EventArgs e)
        {
            string FIOStaff = dataStaff[listBox2.SelectedIndex][0];
            string phoneNumber = dataStaff[listBox2.SelectedIndex][5];
            string subDep = dataStaff[listBox2.SelectedIndex][3];
            string position = dataStaff[listBox2.SelectedIndex][1];
            string rank = dataStaff[listBox2.SelectedIndex][2];
            string shift = dataStaff[listBox2.SelectedIndex][4];
            string date = dateTimePicker1.Value.ToShortDateString();
            string diagnosis = comboBox8.Text;
            string healing = radioButton4.Checked ? "амбулаторное" : "стационарное";
            string vaccinated = checkBox1.Checked ? "+" : "-";

            List<string[]> nameFIO = workingDB.getPatients();

            if (!(nameFIO.Equals(dataStaff[listBox2.SelectedIndex][0])))
            {
                workingDB.addDBPatient(FIOStaff, phoneNumber, subDep, position,  rank,  date,
            diagnosis, healing, shift, vaccinated);

                textBox9.Clear();
                comboBox8.Items.Clear();
                dateTimePicker1.Text = "";
                radioButton3.Checked = false;
                radioButton4.Checked = false;
                checkBox1.Checked = false;

                updateTablePatient();
            }
            else
                MessageBox.Show("Данный пациент уже существует!");   
        }

        //Удаление Пациента из БД
        private void button19_Click(object sender, EventArgs e)
        {
            if (!(comboBox10.SelectedIndex == -1))
            {
                workingDB.deletePatient(dataPatients[comboBox10.SelectedIndex][10]);
                comboBox10.Items.Clear();
                updateTablePatient();
            }
            else
            {
                MessageBox.Show("Не выбран элемент для удаления!");
            }         
        }

        private void tableLayoutPanel5_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}