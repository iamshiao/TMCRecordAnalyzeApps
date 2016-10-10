using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TMC
{
    internal class CL
    {
        public Dictionary<string, DateTime> Proj1 { get; set; } = new Dictionary<string, DateTime>();
        public Dictionary<string, DateTime> Proj2 { get; set; } = new Dictionary<string, DateTime>();
        public Dictionary<string, DateTime> Proj3 { get; set; } = new Dictionary<string, DateTime>();
        public Dictionary<string, DateTime> Proj4 { get; set; } = new Dictionary<string, DateTime>();
        public Dictionary<string, DateTime> Proj5 { get; set; } = new Dictionary<string, DateTime>();
        public Dictionary<string, DateTime> Proj6 { get; set; } = new Dictionary<string, DateTime>();
        public Dictionary<string, DateTime> Proj7 { get; set; } = new Dictionary<string, DateTime>();
        public Dictionary<string, DateTime> Proj8 { get; set; } = new Dictionary<string, DateTime>();
        public Dictionary<string, DateTime> Proj9 { get; set; } = new Dictionary<string, DateTime>();
        public Dictionary<string, DateTime> Proj10 { get; set; } = new Dictionary<string, DateTime>();

        public CL()
        {
            Proj1.Add("AhCounter", DateTime.MinValue);
            Proj1.Add("IE", DateTime.MinValue);
            Proj1.Add("LE", DateTime.MinValue);
            Proj1.Add("TopicsMaster", DateTime.MinValue);

            Proj2.Add("IE", DateTime.MinValue);
            Proj2.Add("LE", DateTime.MinValue);
            Proj2.Add("GE", DateTime.MinValue);

            Proj3.Add("IE", DateTime.MinValue);
            Proj3.Add("LE", DateTime.MinValue);
            Proj3.Add("GE", DateTime.MinValue);

            Proj4.Add("Timer", DateTime.MinValue);
            Proj4.Add("TME", DateTime.MinValue);
            Proj4.Add("Speaker", DateTime.MinValue);
            Proj4.Add("TopicsMaster", DateTime.MinValue);

            Proj5.Add("Speaker", DateTime.MinValue);
            Proj5.Add("GE", DateTime.MinValue);
            Proj5.Add("TME", DateTime.MinValue);
            Proj5.Add("TopicsMaster", DateTime.MinValue);

            Proj7.Add("TME", DateTime.MinValue);
            Proj7.Add("GE", DateTime.MinValue);
            Proj7.Add("Speaker", DateTime.MinValue);

            Proj8.Add("IE", DateTime.MinValue);
            Proj8.Add("GE", DateTime.MinValue);

        }
    }
}
