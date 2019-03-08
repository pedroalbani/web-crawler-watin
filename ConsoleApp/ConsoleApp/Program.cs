using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            Connectors.GoogleScraping scraping = new Connectors.GoogleScraping();
            Console.WriteLine(scraping.FindTerm(args[0]));
        }
    }
}
