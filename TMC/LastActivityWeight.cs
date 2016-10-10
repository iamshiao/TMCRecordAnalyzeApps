using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TMC
{
    internal class LastActivityWeight
    {
        public double Base { get; set; } = 40;
        public double BufferDays { get; set; } = 60;
        public double TME { get; set; } = 0.5;
        public double Timer { get; set; } = 0.8;
        public double AhCounter { get; set; } = 0.8;
        public double Variety { get; set; } = 0.5;
        public double TableTopics { get; set; } = 0.6;
        public double GE { get; set; } = 0.5;
        public double IE { get; set; } = 0.7;
        public double LE { get; set; } = 0.7;
        public double Speech { get; set; } = 0.35;
    }
}
