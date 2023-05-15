using Button = System.Windows.Forms.Button;
using Color = System.Drawing.Color;

namespace MCHSAutoTable
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            _dataEdds = _workingDb.GetEdds();
        }

        //Статус кнопок для изменения цвета и закрытия панелей
        byte _btn1Status = 0;
        byte _btn2Status = 0;
        byte _btn3Status = 0;
        byte _btn4Status = 0;

        int _indexDataEdds = 0;

        //Для работы с БД
        WorkingDb _workingDb = new WorkingDb();

        //Вывод информации из БД в WORD или Excel
        InputDataWordExcel _inputData = new InputDataWordExcel();

        //ЕДДС
        private void button1_Click(object sender, EventArgs e)
        {
            //inputData.createTableExcelPatients();
            _workingDb.ClearRowInTableEdds();
            listBox1.Items.Clear();

            //InputDataWordExcel inputDataWordExcel = new InputDataWordExcel();
            //inputDataWordExcel.createTableExcelPatients();

            //Загрузка в список
            if (_dataEdds.Count > 0)
            {
                _dataEdds = _workingDb.GetEdds();
                label3.Text = "Номер телефона: " + _dataEdds[_indexDataEdds][1];
                foreach (string[] s in _dataEdds)
                {
                    listBox1.Items.Add(s[0]);
                }

                radioButton1.Checked = true;
                listBox1.SetSelected(0, true);

                _btn1Status = ButtonColorAndOpenPanel(button1, _btn1Status, panel1);
            }
            else
                MessageBox.Show("Заполните данные в разделе Ведомства ЕДДС");
        }

        //Времено нетрудоспособные
        private void button2_Click(object sender, EventArgs e)
        {
            listBox2.Items.Clear();
            UpdateDataStaff();
            UpdateDataDiagnosis();
            UpdateTablePatient();
            _btn2Status = ButtonColorAndOpenPanel(button2, _btn2Status, panel6);
        }

        //Ведомства ЕДДС
        private void button4_Click(object sender, EventArgs e)
        {
            _btn3Status = ButtonColorAndOpenPanel(button4, _btn3Status, panel2);

            //Обновление БД
            UpdateTableEdds();
        }

        //Работники
        private void button6_Click(object sender, EventArgs e)
        {
            comboBox2.Items.Clear();
            comboBox4.Items.Clear();
            comboBox5.Items.Clear();

            UpdateDataPosition();
            UpdateDataSubDepartment();
            UpdateDataStaff();

            _btn4Status = ButtonColorAndOpenPanel(button6, _btn4Status, panel3);
            List<String> listRank = new List<String>()
            {
                "рядовой", "мл. сержант", "сержант", "ст. сержант", "старшина",
                "прапорщик", "ст. прапорщик", "мл. лейтенант", "лейтенант", "ст. лейтенант", "капитан", "майор",
                "подполковник", "полковник"
            };

            List<String> listShift = new List<String>() { "-", "1", "2", "3" };

            comboBox6.DataSource = listShift;
            comboBox3.DataSource = listRank;
        }

        //Обновление таблицы ведомства ЕДДС
        List<string[]> _dataEdds = new List<string[]>();

        void UpdateTableEdds()
        {
            dataGridView1.Rows.Clear();
            _dataEdds.Clear();
            comboBox1.Items.Clear();
            comboBox1.ResetText();
            textBox4.Clear();
            maskedTextBox1.Clear();

            _dataEdds = _workingDb.GetEdds();

            if (_dataEdds.Count > 0)
            {
                foreach (string[] s in _dataEdds)
                {
                    dataGridView1.Rows.Add(s);
                    comboBox1.Items.Add(s[0]);
                }
            }
        }

        //Обновление таблицы ведомства ЕДДС
        List<string[]> _dataPatients = new List<string[]>();

        void UpdateTablePatient()
        {
            dataGridView3.Rows.Clear();
            comboBox10.Items.Clear();

            _dataPatients = _workingDb.GetPatients();

            if (_dataPatients.Count > 0)
            {
                foreach (string[] s in _dataPatients)
                {
                    dataGridView3.Rows.Add(s);
                    comboBox10.Items.Add(s[0]);
                }
            }
        }

        //Обновление checkBox с должностями
        List<string[]> _dataPosition = new List<string[]>();

        void UpdateDataPosition()
        {
            _dataPosition = _workingDb.GetPosition();

            comboBox2.Items.Clear();
            comboBox7.Items.Clear();

            if (_dataPosition.Count > 0)
            {
                foreach (string[] s in _dataPosition)
                {
                    comboBox2.Items.Add(s[0]);
                    comboBox7.Items.Add(s[0]);
                }
            }
        }

        //Обновление checkBox с подразделениями
        List<string[]> _dataSubDep = new List<string[]>();

        void UpdateDataSubDepartment()
        {
            _dataSubDep = _workingDb.GetSubDepartment();

            comboBox4.Items.Clear();
            comboBox7.Items.Clear();

            if (_dataSubDep.Count > 0)
            {
                foreach (string[] s in _dataSubDep)
                {
                    comboBox4.Items.Add(s[0]);
                    comboBox7.Items.Add(s[0]);
                }
            }
        }

        //Обновление checkBox с сотрудниками
        List<string[]> _dataStaff = new List<string[]>();

        void UpdateDataStaff()
        {
            dataGridView2.Rows.Clear();
            comboBox5.Items.Clear();
            _dataStaff = _workingDb.GetStaff();

            if (_dataStaff.Count > 0)
            {
                foreach (string[] s in _dataStaff)
                {
                    dataGridView2.Rows.Add(s);
                    comboBox5.Items.Add(s[0]);
                    listBox2.Items.Add(s[0]);
                }
            }
        }

        //Обновление checkBox с Болезнями
        List<string[]> _dataDiagnosis = new List<string[]>();

        void UpdateDataDiagnosis()
        {
            _dataDiagnosis = _workingDb.GetDiagnosis();

            comboBox8.Items.Clear();
            comboBox9.Items.Clear();

            if (_dataDiagnosis.Count > 0)
            {
                foreach (string[] s in _dataDiagnosis)
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
            _workingDb.AddDbDepartmentEdds(name, phoneNumber);
            UpdateTableEdds();
        }

        //Удалить ведомство ЕДДС
        private void button7_Click(object sender, EventArgs e)
        {
            if (!(comboBox1.SelectedIndex == -1))
            {
                _workingDb.DeleteEdds(_dataEdds[comboBox1.SelectedIndex][2]);
                UpdateTableEdds();
            }
            else
            {
                MessageBox.Show("Не выбран элемент для удаления!");
            }
        }

        List<string[]> _tableDataEdds = new List<string[]>();

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
                label3.Text = "Номер телефона: " + _dataEdds[_indexDataEdds][1];
                _workingDb.AddDbTableEdds(fedds, time, working);

                textBox2.Clear();

                if (_dataEdds.Count() - 1 != listBox1.SelectedIndex)
                {
                    listBox1.SelectedIndex += 1;
                }
                else
                {
                    _btn1Status = 0;
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
            _tableDataEdds = _workingDb.GetTableEdds();
            //Вывод в WORD
            InputDataWordExcel.CreateTableWordEdds(textBox6.Text, _tableDataEdds);
            panel1.Hide();
            button3.Enabled = true;
            panel5.Visible = false;
            textBox2.Enabled = true;
            textBox6.Clear();
        }

        //Смена цвета для кнопок МЕНЮ
        byte ButtonColorAndOpenPanel(Button button, byte btn, Panel panel)
        {
            if (btn == 0)
            {
                InvisibleAllPanel();
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

            string phoneNumberTable = _dataEdds[listBox1.SelectedIndex][1];
            label3.Text = "Номер телефона: " + phoneNumberTable;
        }

        //Скрытие всех панелей и кнопок МЕНЮ
        void InvisibleAllPanel()
        {
            panel1.Visible = false;
            panel2.Visible = false;
            panel3.Visible = false;
            panel6.Visible = false;

            button1.BackColor = Color.ForestGreen;
            button2.BackColor = Color.ForestGreen;
            button4.BackColor = Color.ForestGreen;
            button6.BackColor = Color.ForestGreen;

            _btn1Status = 0;
            _btn2Status = 0;
            _btn3Status = 0;
            _btn4Status = 0;
        }

        int _indexDataDbWorkStaff = 0;

        //Смена информации в Панели в разделе Работники
        public void MessageBoxAddData(int paramSwitch)
        {
            switch (paramSwitch)
            {
                case 0:
                    label15.Text = "Добавление должности";
                    _indexDataDbWorkStaff = 1;
                    UpdateDataPosition();
                    panel4.Visible = true;
                    break;
                case 1:
                    label15.Text = "Добавление подразделения";
                    _indexDataDbWorkStaff = 0;
                    UpdateDataSubDepartment();
                    panel4.Visible = true;
                    break;
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            MessageBoxAddData(0);
        }

        private void button10_Click(object sender, EventArgs e)
        {
            MessageBoxAddData(1);
        }

        private void button12_Click(object sender, EventArgs e)
        {
            if (_indexDataDbWorkStaff == 1)
            {
                _workingDb.AddDbPosition(textBox5.Text);
                comboBox2.Items.Clear();
                UpdateDataPosition();
            }
            else
            {
                _workingDb.AddDbSubDepartment(textBox5.Text);
                comboBox4.Items.Clear();
                UpdateDataSubDepartment();
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
            if (!(textBox3.Text.Equals("") || comboBox2.Text.Equals("") || comboBox4.Text.Equals("") ||
                  maskedTextBox2.Text.Equals("") || comboBox3.Text.Equals("") || comboBox6.Text.Equals("")))
            {
                _workingDb.AddDbStaff(textBox3.Text, comboBox2.Text, comboBox4.Text, maskedTextBox2.Text,
                    comboBox3.Text, comboBox6.Text);

                textBox3.Clear();
                comboBox2.Text = "";
                comboBox4.Text = "";
                maskedTextBox2.Text = "";
                comboBox3.Text = "";
                comboBox6.Text = "";

                UpdateDataStaff();
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
                _workingDb.DeleteStaff(_dataStaff[comboBox5.SelectedIndex][6]);
                UpdateDataStaff();
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
                if (_indexDataDbWorkStaff == 1)
                {
                    _workingDb.DeletePosition(_dataPosition[comboBox7.SelectedIndex][1]);
                    UpdateDataPosition();
                }
                else
                {
                    _workingDb.DeleteSubDep(_dataSubDep[comboBox7.SelectedIndex][1]);
                    UpdateDataSubDepartment();
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
            UpdateDataDiagnosis();
            panel7.Visible = true;
        }

        //Добавление Диагноза в БД
        private void button18_Click(object sender, EventArgs e)
        {
            _workingDb.AddDbDiagnosis(textBox7.Text);
            UpdateDataDiagnosis();

            textBox7.Clear();
            comboBox9.Items.Clear();
            panel7.Visible = false;
        }

        //Удаление диагноза из БД
        private void button17_Click(object sender, EventArgs e)
        {
            if (!(comboBox9.SelectedIndex == -1))
            {
                _workingDb.DeleteDiagnosis(_dataDiagnosis[comboBox9.SelectedIndex][1]);
                UpdateDataDiagnosis();
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
            if (_dataStaff != null && _dataStaff.Count != 0)
            {
                string fioStaff = _dataStaff[listBox2.SelectedIndex][0];
                string phoneNumber = _dataStaff[listBox2.SelectedIndex][5];
                string subDep = _dataStaff[listBox2.SelectedIndex][3];
                string position = _dataStaff[listBox2.SelectedIndex][1];
                string rank = _dataStaff[listBox2.SelectedIndex][2];
                string shift = _dataStaff[listBox2.SelectedIndex][4];
                string date = dateTimePicker1.Value.ToShortDateString();
                string diagnosis = comboBox8.Text;
                string healing = radioButton4.Checked ? "амбулаторное" : "стационарное";
                string vaccinated = checkBox1.Checked ? "+" : "-";

                List<string[]> nameFio = _workingDb.GetPatients();


                if (!(nameFio.Equals(_dataStaff[listBox2.SelectedIndex][0])))
                {
                    _workingDb.AddDbPatient(fioStaff, phoneNumber, subDep, position, rank, date,
                        diagnosis, healing, shift, vaccinated);

                    textBox9.Clear();
                    comboBox8.Items.Clear();
                    dateTimePicker1.Text = "";
                    radioButton3.Checked = false;
                    radioButton4.Checked = false;
                    checkBox1.Checked = false;

                    UpdateTablePatient();
                }
                else
                    MessageBox.Show("Данный пациент уже существует!");
            }
            else MessageBox.Show("Не выбран сотрудник!");
        }

        //Удаление Пациента из БД
            private void button19_Click(object sender, EventArgs e)
            {
                if (!(comboBox10.SelectedIndex == -1))
                {
                    _workingDb.DeletePatient(_dataPatients[comboBox10.SelectedIndex][10]);
                    comboBox10.Items.Clear();
                    UpdateTablePatient();
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