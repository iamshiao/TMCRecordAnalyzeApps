using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TMC
{
    internal class MyIERecord
    {
        public string Name { get; set; }

        public List<IERecord> Records { get; set; }
    }

    internal class IERecord
    {
        public string ProjLevel { get; set; }

        public string Name { get; set; }

        public DateTime MeetingDate { get; set; }
    }
}
