using System;
using System.Collections.Generic;
using System.Diagnostics;
using AnalisWordLib;

namespace testApp
{
    class Program
    {

        static void Main()
        {
            Stopwatch stopwatch = new Stopwatch();
            DataUnit unit = new DataUnit();
            Console.WriteLine("Start.");
            stopwatch.Start();
            AnalisDocument analis = new AnalisDoc(String.Empty);
            try
            {
                unit = analis.Parse("D:\\test\\План-график Поляков.doc");
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            
            stopwatch.Stop();
            Console.WriteLine(unit.Owner);
            Console.WriteLine("End. Total time: " + stopwatch.Elapsed);
            Console.ReadKey();
        }
    }
}
