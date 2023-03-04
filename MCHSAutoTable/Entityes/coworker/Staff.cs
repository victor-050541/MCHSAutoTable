using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MCHSAutoTable.Entityes.coworker
{
    public class Staff
    {
        public int StaffId { get; set; }
        public string FIO { get; set; }
        public string PositionName { get; set; }
        public string Rank { get; set; }
        public string SubDepatment { get; set; }
        public string Shift { get; set; }
        public string PhoneNumber { get; set; }

    
        //public int PositionId { get; set; }
        //public virtual Position Position { get; set; }
        //public int RankId { get; set; }
        //public virtual Rank Rank { get; set; }

        //public virtual SubDepartment SubDepartmen { get; set; }        
    }
}
