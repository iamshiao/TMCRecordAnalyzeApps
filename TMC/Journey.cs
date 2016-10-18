using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TMC
{
    internal class Journey
    {
        public string Name { get; set; }
        public List<Role> Achievements { get; set; }

    }

    internal class Role
    {
        public string Name { get; set; }
        public DateTime Date { get; set; }
    }
}
