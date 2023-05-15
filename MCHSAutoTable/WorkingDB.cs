using MCHSAutoTable.Entities.Coworker;
using MCHSAutoTable.Entities.EDDS;


namespace MCHSAutoTable
{
    public class WorkingDb
    {
        //Добавление пациента в заболевшие
        public void AddDbPatient(
            string fioStaff,
            string phoneNumber,
            string subDep,
            string position,
            string rank,
            string date,
            string diagnosis,
            string healing,
            string shift,
            string vaccinated
        )
        {
            using (ApplicationContext db = new ApplicationContext())
            {
                db.Patients.Add(new Patients
                {
                    Fio = fioStaff, PhoneNumber = phoneNumber, SubDepartment = subDep, Position = position, Rank = rank,
                    Date = date, Diagnosis = diagnosis, Shift = shift, Healing = healing, Vaccinated = vaccinated
                });
                db.SaveChanges();
            }
        }

        //Добавление Диагноза
        public void AddDbDiagnosis(string name)
        {
            using (ApplicationContext db = new ApplicationContext())
            {
                db.Diagnoses.Add(new Diagnosis { Name = name });
                db.SaveChanges();
            }
        }

        //Добавление обзвона ЕДДС по таблице
        public void AddDbTableEdds(string fio, string time, string working)
        {
            using (ApplicationContext db = new ApplicationContext())
            {
                db.TableEdds.Add(new TableEdds { Fio = fio, Time = time, Working = working });
                db.SaveChanges();
            }
        }

        //Добавление ЕДДС
        public void AddDbDepartmentEdds(string name, string phoneNumber)
        {
            using (ApplicationContext db = new ApplicationContext())
            {
                db.Edds.Add(new Edds { Name = name, PhoneNumber = phoneNumber });
                db.SaveChanges();
            }
        }

        //Добавление сотрудника    
        public void AddDbStaff(
            string fio,
            string position,
            string subDep,
            string phoneNumber,
            string rank,
            string shift
        )
        {
            using (ApplicationContext db = new ApplicationContext())
            {
                db.Staffs.Add(new Staff
                {
                    Fio = fio, PositionName = position, SubDepatment = subDep, Shift = shift, PhoneNumber = phoneNumber,
                    Rank = rank
                });
                db.SaveChanges();
            }
        }

        //Добавление должонсти
        public void AddDbPosition(string name)
        {
            using (ApplicationContext db = new ApplicationContext())
            {
                db.Positions.Add(new Position { Name = name });
                db.SaveChanges();
            }
        }

        //Добавление подразделения
        public void AddDbSubDepartment(string name)
        {
            using (ApplicationContext db = new ApplicationContext())
            {
                db.SubDepartments.Add(new SubDepartment { Name = name });
                db.SaveChanges();
            }
        }

        //Получение элементов ведомства ЕДДС из БД
        public List<String[]> GetEdds()
        {
            List<String[]> dataEdds = new List<String[]>();
            int i = 0;

            using (ApplicationContext db = new ApplicationContext())
            {
                var edds = db.Edds.ToList();
                foreach (Edds e in edds)
                {
                    dataEdds.Add(new string[3]);
                    dataEdds[i][0] = e.Name;
                    dataEdds[i][1] = e.PhoneNumber;
                    dataEdds[i][2] = Convert.ToString(e.EddsId);
                    i++;
                }

                return dataEdds;
            }
        }

        //Получение данных из таблицы TableEDDS для вывода в WORD
        public List<String[]> GetTableEdds()
        {
            List<String[]> dataEdds = new List<String[]>();
            int i = 0;

            using (ApplicationContext db = new ApplicationContext())
            {
                var edds = db.TableEdds.ToList();
                foreach (TableEdds e in edds)
                {
                    dataEdds.Add(new string[5]);
                    dataEdds[i][0] = e.Time;
                    dataEdds[i][1] = GetEdds()[i][0];
                    dataEdds[i][2] = e.Working;
                    dataEdds[i][3] = e.Fio;
                    dataEdds[i][4] = GetEdds()[i][1];
                    ;

                    i++;
                }

                return dataEdds;
            }
        }

        //Получение данных из таблицы Position
        public List<String[]> GetPosition()
        {
            List<String[]> dataPosition = new List<String[]>();
            int i = 0;

            using (ApplicationContext db = new ApplicationContext())
            {
                var pos = db.Positions.ToList();
                foreach (Position e in pos)
                {
                    dataPosition.Add(new string[2]);
                    dataPosition[i][0] = e.Name;
                    dataPosition[i][1] = Convert.ToString(e.PositionId);
                    i++;
                }

                return dataPosition;
            }
        }

        //Получение данных из таблицы SubDepatment
        public List<String[]> GetSubDepartment()
        {
            List<String[]> dataSubDepartment = new List<String[]>();
            int i = 0;

            using (ApplicationContext db = new ApplicationContext())
            {
                var subDep = db.SubDepartments.ToList();
                foreach (SubDepartment e in subDep)
                {
                    dataSubDepartment.Add(new string[2]);
                    dataSubDepartment[i][0] = e.Name;
                    dataSubDepartment[i][1] = Convert.ToString(e.SubDepartmentId);
                    i++;
                }

                return dataSubDepartment;
            }
        }

        //Получение данных из таблицы SubDepatment
        public List<String[]> GetDiagnosis()
        {
            List<String[]> dataDiagnosis = new List<String[]>();
            int i = 0;

            using (ApplicationContext db = new ApplicationContext())
            {
                var diagnosis = db.Diagnoses.ToList();
                foreach (Diagnosis e in diagnosis)
                {
                    dataDiagnosis.Add(new string[2]);
                    dataDiagnosis[i][0] = e.Name;
                    dataDiagnosis[i][1] = Convert.ToString(e.DiagnosisId);
                    i++;
                }

                return dataDiagnosis;
            }
        }

        //Получение данных из таблицы Staff
        public List<String[]> GetStaff()
        {
            List<String[]> dataStaff = new List<String[]>();
            int i = 0;

            using (ApplicationContext db = new ApplicationContext())
            {
                var staff = db.Staffs.ToList();
                foreach (Staff e in staff)
                {
                    dataStaff.Add(new string[7]);
                    dataStaff[i][0] = e.Fio;
                    dataStaff[i][1] = e.PositionName;
                    dataStaff[i][2] = e.Rank;
                    dataStaff[i][3] = e.SubDepatment;
                    dataStaff[i][4] = e.Shift;
                    dataStaff[i][5] = e.PhoneNumber;
                    dataStaff[i][6] = Convert.ToString(e.StaffId);
                    i++;
                }

                return dataStaff;
            }
        }

        //Получение данных из таблицы Patient
        public List<String[]> GetPatients()
        {
            List<String[]> dataPatient = new List<String[]>();
            int i = 0;

            using (ApplicationContext db = new ApplicationContext())
            {
                var patients = db.Patients.ToList();
                foreach (Patients e in patients)
                {
                    dataPatient.Add(new string[11]);
                    dataPatient[i][0] = e.Fio;
                    dataPatient[i][1] = e.Date;
                    dataPatient[i][2] = e.Diagnosis;
                    dataPatient[i][3] = e.PhoneNumber;
                    dataPatient[i][4] = e.SubDepartment;
                    dataPatient[i][5] = e.Position;
                    dataPatient[i][6] = e.Rank;
                    dataPatient[i][7] = e.Shift;
                    dataPatient[i][8] = e.Healing;
                    dataPatient[i][9] = e.Vaccinated;
                    dataPatient[i][10] = Convert.ToString(e.PatientsId);
                    i++;
                }

                return dataPatient;
            }
        }

        //Удаление всех записей в таблице TableEDDS
        public void ClearRowInTableEdds()
        {
            using (ApplicationContext db = new ApplicationContext())
            {
                db.TableEdds.RemoveRange(db.TableEdds);
                db.SaveChanges();
            }
        }


        //Удаление ведомство ЕДДС из БД
        public void DeleteEdds(string idEdds)
        {
            using (ApplicationContext db = new ApplicationContext())
            {
                Edds edds = new Edds() { EddsId = Convert.ToInt32(idEdds) };
                db.Edds.Attach(edds);
                db.Edds.Remove(edds);
                db.SaveChanges();
            }
        }

        //Удаление сотрудника из БД
        public void DeleteStaff(string idStaff)
        {
            using (ApplicationContext db = new ApplicationContext())
            {
                Staff staff = new Staff() { StaffId = Convert.ToInt32(idStaff) };
                db.Staffs.Attach(staff);
                db.Staffs.Remove(staff);
                db.SaveChanges();
            }
        }

        //Удаление подразделения из БД
        public void DeleteSubDep(string idDep)
        {
            using (ApplicationContext db = new ApplicationContext())
            {
                SubDepartment subDep = new SubDepartment() { SubDepartmentId = Convert.ToInt32(idDep) };
                db.SubDepartments.Attach(subDep);
                db.SubDepartments.Remove(subDep);
                db.SaveChanges();
            }
        }

        //Удаление должности из БД
        public void DeletePosition(string idPos)
        {
            using (ApplicationContext db = new ApplicationContext())
            {
                Position pos = new Position() { PositionId = Convert.ToInt32(idPos) };
                db.Positions.Attach(pos);
                db.Positions.Remove(pos);
                db.SaveChanges();
            }
        }

        //Удаление заболевания из БД
        public void DeleteDiagnosis(string idDiagnosis)
        {
            using (ApplicationContext db = new ApplicationContext())
            {
                Diagnosis diagnosis = new Diagnosis() { DiagnosisId = Convert.ToInt32(idDiagnosis) };
                db.Diagnoses.Attach(diagnosis);
                db.Diagnoses.Remove(diagnosis);
                db.SaveChanges();
            }
        }

        //Удаление пациента из БД
        public void DeletePatient(string idPatient)
        {
            using (ApplicationContext db = new ApplicationContext())
            {
                Patients patients = new Patients() { PatientsId = Convert.ToInt32(idPatient) };
                db.Patients.Attach(patients);
                db.Patients.Remove(patients);
                db.SaveChanges();
            }
        }
    }
}