using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TMC
{
    internal class RecentActivity
    {
        public string Name { get; set; }
        public int TME { get; set; }
        public int Timer { get; set; }
        public int AhCounter { get; set; }
        public int Variety { get; set; }
        public int TableTopics { get; set; }
        public int GE { get; set; }
        public int IE { get; set; }
        public int LE { get; set; }
        public double HistoryCredit { get; set; }
        public string LastSpeech { get; set; }
        public DateTime LastSpeechAt { get; set; }
        public double CurrSpeechCredit { get; set; }
        public string LastAssignment { get; set; }
        public DateTime LastAssignAt { get; set; }
        public double CurrAssignCredit { get; set; }
    }
}
