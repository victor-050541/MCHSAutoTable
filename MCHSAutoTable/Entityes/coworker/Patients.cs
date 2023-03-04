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
        public string StaffId { get; set; }
        public string Date { get; set; }
        public string Diagnosis { get; set; }
        public string Healing { get; set; }
        public string Vaccinated { get; set; }
    }
}
