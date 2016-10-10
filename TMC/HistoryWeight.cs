using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TMC
{
    internal class HistoryWeight
    {
        public double Base { get; set; } = 20;
        public double TME { get; set; } = 0.2;
        public double Timer { get; set; } = 0.05;
        public double AhCounter { get; set; } = 0.05;
        public double Variety { get; set; } = 0.2;
        public double TableTopics { get; set; } = 0.15;
        public double GE { get; set; } = 0.15;
        public double IE { get; set; } = 0.1;
        public double LE { get; set; } = 0.1;
    }
}
