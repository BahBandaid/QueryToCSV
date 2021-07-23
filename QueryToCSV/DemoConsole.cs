using System;

namespace QueryToCSV
{
    internal class DemoConsole
    {
        private static void Main()
        {
            Console.WriteLine("Enter Connection String:");
            var connString = Console.ReadLine();
            Console.WriteLine("Enter Query String:");
            var queryString = Console.ReadLine();
            Console.WriteLine("Enter Full CSV File Path:");
            var csvPath = Console.ReadLine();
            QueryCSVHelper.OutputCsv(connString, queryString, csvPath);

        }
    }
}
