using DocumentFormat.OpenXml.ExtendedProperties;
using MCHSAutoTable.Entityes.coworker;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Numerics;
using System.Text;
using System.Threading.Tasks;

namespace MCHSAutoTable.Entityes.edds
{
    public class TableEDDS
    {
        public int TableEDDSId { get; set; }
        public string Time { get; set; }
        public int EDDSId { get; set; }
        public string FIO { get; set; }
              
        public string Working { get; set; }
        
    }
}
