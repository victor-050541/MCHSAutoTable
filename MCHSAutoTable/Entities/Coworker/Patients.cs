using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MCHSAutoTable.Entityes.coworker
{
    public class Patients
    {
        public int PatientsId { get; set; }
        public string FIO { get; set; }
        public string PhoneNumber { get; set; }
        public string SubDepartment { get; set; }
        public string Position { get; set; }
        public string Rank { get; set; }
        public string Date { get; set; }
        public string Diagnosis { get; set; }
        public string Shift { get; set; }
        public string Healing { get; set; }
        public string Vaccinated { get; set; }
    }
}
