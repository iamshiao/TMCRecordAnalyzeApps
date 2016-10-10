using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace NamesOrder
{
    class Program
    {
        static void Main(string[] args)
        {
            if (File.Exists(@"names.txt")) {
                List<string> nameColl = File.ReadAllLines(@"names.txt").Select(name => name.Trim()).ToList();

                var ret = nameColl.Distinct().OrderBy(s => s).ToList();
                File.WriteAllLines(@"names.txt", ret);
                Console.WriteLine("Reorder succeed.");
            }
            else {
                Console.WriteLine("Can't find names.txt file.");
            }

            Console.WriteLine("Prepare self close.");
            Thread.Sleep(1500);
        }
    }
}
