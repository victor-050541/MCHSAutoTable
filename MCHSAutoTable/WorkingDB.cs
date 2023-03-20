using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.InkML;
using DocumentFormat.OpenXml.Office2010.Excel;
using MCHSAutoTable.Entityes.coworker;
using MCHSAutoTable.Entityes.edds;
using Microsoft.EntityFrameworkCore;

namespace MCHSAutoTable
{
    public class WorkingDB
    {
        //Добавление пациента в заболевшие
        public void addDBPatient(string FIOStaff, string PhoneNumber, string SubDep, string Position, string Rank, string Date,
            string Diagnosis, string Healing, string Shift, string Vaccinated)
        {
            using (ApplicationContext db = new ApplicationContext())
            {                
                db.Patients.Add(new Patients { FIO = FIOStaff, PhoneNumber = PhoneNumber, SubDepartment = SubDep, Position = Position, Rank = Rank,
                    Date = Date, Diagnosis = Diagnosis, Shift = Shift, Healing = Healing, Vaccinated = Vaccinated });                
                db.SaveChanges();                                
            }
        }

        //Добавление Диагноза
        public void addDBDiagnosis(string Name)
        {
            using (ApplicationContext db = new ApplicationContext())
            {
                db.Diagnoses.Add(new Diagnosis { Name = Name });
                db.SaveChanges();
            }
        }

        //Добавление обзвона ЕДДС по таблице
        public void addDBTableEDDS(string FIO, string Time, string Working)
        {
            using (ApplicationContext db = new ApplicationContext())
            {
                
                db.TableEDDS.Add(new TableEDDS { FIO = FIO, Time = Time, Working = Working });                
                db.SaveChanges();
            }
        }

        //Добавление ЕДДС
        public void addDBDepartmentEDDS(string Name, string PhoneNumber)
        {
            using (ApplicationContext db = new ApplicationContext())
            {
                db.EDDS.Add(new EDDS { Name = Name, PhoneNumber = PhoneNumber});
                db.SaveChanges();
            }
        }

        //Добавление сотрудника    
        public void addDBStaff(string FIO, string Position, string SubDep,string PhoneNumber,string Rank, string Shift)
        {
            using (ApplicationContext db = new ApplicationContext())
            {
                db.Staffs.Add(new Staff { FIO = FIO, PositionName = Position, SubDepatment = SubDep, Shift = Shift, PhoneNumber = PhoneNumber, Rank = Rank});
                db.SaveChanges();
            }
        }

        //Добавление должонсти
        public void addDBPosition(string Name)
        {
            using (ApplicationContext db = new ApplicationContext())
            {
                db.Positions.Add(new Position { Name = Name});
                db.SaveChanges();
            }
        }

        //Добавление подразделения
        public void addDBSubDepartment(string Name)
        {
            using (ApplicationContext db = new ApplicationContext())
            {
                db.SubDepartments.Add(new SubDepartment { Name = Name });
                db.SaveChanges();
            }
        }

        //Получение элементов ведомства ЕДДС из БД
        public List<String[]> getEDDS()
        {
            List<String[]> dataEDDS = new List<String[]>();
            int i = 0;

            using (ApplicationContext db = new ApplicationContext())
            {
                var edds = db.EDDS.ToList();
                foreach (EDDS e in edds)
                {
                    dataEDDS.Add(new string[3]);
                    dataEDDS[i][0] = e.Name;
                    dataEDDS[i][1] = e.PhoneNumber;
                    dataEDDS[i][2] = Convert.ToString(e.EDDSId);
                    i++;
                }
                return dataEDDS;
            }            
        }

        //Получение данных из таблицы TableEDDS для вывода в WORD
        public List<String[]> getTableEDDS()
        {
            List<String[]> dataEDDS = new List<String[]>();
            int i = 0;

            using (ApplicationContext db = new ApplicationContext())
            {
                var edds = db.TableEDDS.ToList();
                foreach (TableEDDS e in edds)
                {
                    dataEDDS.Add(new string[5]);
                    dataEDDS[i][0] = e.Time;
                    dataEDDS[i][1] = getEDDS()[i][0];
                    dataEDDS[i][2] = e.Working;
                    dataEDDS[i][3] = e.FIO;
                    dataEDDS[i][4] = getEDDS()[i][1]; ;
                    
                    i++;
                }
                return dataEDDS;
            }
        }

        //Получение данных из таблицы Position
        public List<String[]> getPosition()
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
        public List<String[]> getSubDepartment()
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
        public List<String[]> getDiagnosis()
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
        public List<String[]> getStaff()
        {
            List<String[]> dataStaff = new List<String[]>();
            int i = 0;

            using (ApplicationContext db = new ApplicationContext())
            {
                var staff = db.Staffs.ToList();
                foreach (Staff e in staff)
                {
                    dataStaff.Add(new string[7]);
                    dataStaff[i][0] = e.FIO;
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
        public List<String[]> getPatients()
        {
            List<String[]> dataPatient = new List<String[]>();
            int i = 0;

            using (ApplicationContext db = new ApplicationContext())
            {
                var patients = db.Patients.ToList();
                foreach (Patients e in patients)
                {
                    dataPatient.Add(new string[11]);
                    dataPatient[i][0] = e.FIO;
                    dataPatient[i][1] = e.PhoneNumber;
                    dataPatient[i][2] = e.SubDepartment;
                    dataPatient[i][3] = e.Position;
                    dataPatient[i][4] = e.Rank;
                    dataPatient[i][5] = e.Date;
                    dataPatient[i][7] = e.Diagnosis;
                    dataPatient[i][6] = e.Shift;
                    dataPatient[i][8] = e.Healing;
                    dataPatient[i][9] = e.Vaccinated;
                    dataPatient[i][10] = Convert.ToString(e.PatientsId);
                    i++;
                }
                return dataPatient;
            }
        }

        //Удаление всех записей в таблице TableEDDS
        public void clearRowInTableEDDS()
        {
            using (ApplicationContext db = new ApplicationContext())
            {
                db.TableEDDS.RemoveRange(db.TableEDDS);
                db.SaveChanges();
            }
        }

        
        //Удаление ведомство ЕДДС из БД
        public void deleteEDDS(string IdEDDS)
        {
            using (ApplicationContext db = new ApplicationContext())
            {                
                EDDS edds = new EDDS() { EDDSId = Convert.ToInt32(IdEDDS)};
                db.EDDS.Attach(edds);
                db.EDDS.Remove(edds);
                db.SaveChanges();
            }
        }

        //Удаление сотрудника из БД
        public void deleteStaff(string IdStaff)
        {
            using (ApplicationContext db = new ApplicationContext())
            {
                Staff staff = new Staff() { StaffId = Convert.ToInt32(IdStaff) };
                db.Staffs.Attach(staff);
                db.Staffs.Remove(staff);
                db.SaveChanges();
            }
        }

        //Удаление подразделения из БД
        public void deleteSubDep(string IdDep)
        {
            using (ApplicationContext db = new ApplicationContext())
            {
                SubDepartment subDep = new SubDepartment() { SubDepartmentId = Convert.ToInt32(IdDep) };
                db.SubDepartments.Attach(subDep);
                db.SubDepartments.Remove(subDep);
                db.SaveChanges();
            }
        }

        //Удаление должности из БД
        public void deletePosition(string IdPos)
        {
            using (ApplicationContext db = new ApplicationContext())
            {
                Position pos = new Position() { PositionId = Convert.ToInt32(IdPos) };
                db.Positions.Attach(pos);
                db.Positions.Remove(pos);
                db.SaveChanges();
            }
        }

        //Удаление заболевания из БД
        public void deleteDiagnosis(string IdDiagnosis)
        {
            using (ApplicationContext db = new ApplicationContext())
            {
                Diagnosis diagnosis = new Diagnosis() { DiagnosisId = Convert.ToInt32(IdDiagnosis) };
                db.Diagnoses.Attach(diagnosis);
                db.Diagnoses.Remove(diagnosis);
                db.SaveChanges();
            }
        }

        //Удаление пациента из БД
        public void deletePatient(string IdPatient)
        {
            using (ApplicationContext db = new ApplicationContext())
            {
                Patients patients = new Patients() { PatientsId = Convert.ToInt32(IdPatient) };
                db.Patients.Attach(patients);
                db.Patients.Remove(patients);
                db.SaveChanges();
            }
        }
    }
}
