using CsvHelper.Configuration;

namespace RCSVB.Models
{
    public sealed class RealmsRecordMap : ClassMap<RealmsRecord>
    {
        public RealmsRecordMap()
        {
            Map(m => m.Account).Index(0);
            Map(m => m.Actual).Index(1);
            Map(m => m.Budget).Index(2);
            Map(m => m.Variance).Index(3);
        }
    }
}
