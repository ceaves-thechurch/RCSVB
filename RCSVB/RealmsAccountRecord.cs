using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CsvHelper.Configuration;

namespace RCSVB
{
    public class RealmsAccountRecord
    {
        public string Owner { get; set; }
        public string Department { get; set; }
        public string Account { get; set; }
        public string Actual { get; set; }
        public string Budget { get; set; }
        public string Variance { get; set; }

    }

    public sealed class RealmsAccountRecordMap : ClassMap<RealmsAccountRecord>
    {
        public RealmsAccountRecordMap ()
        {
            Map(m => m.Account).Index(0);
            Map(m => m.Actual).Index(1);
            Map(m => m.Budget).Index(2);
            Map(m => m.Variance).Index(3);
        }
    }
}
