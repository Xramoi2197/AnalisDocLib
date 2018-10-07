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
            List<string> Docs = new List<string>();
            Docs.Add("D:\\test\\План-график Козуб.doc");
            Docs.Add("D:\\test\\План-график Коз.doc");
            Docs.Add("D:\\test\\План-график Поляков.doc");
            Docs.Add("D:\\test\\План-график Козуб.doc");
            Docs.Add("D:\\test\\Word.doc");
            Model newModel = new Model(String.Empty, "XML");
            foreach (var doc in Docs)
            {
                try
                {
                    
                    newModel.ParseDoc(doc);
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                }
            }
                     
            stopwatch.Stop();
            Console.WriteLine(unit.Owner);
            Console.WriteLine("End. Total time: " + stopwatch.Elapsed);
            Console.ReadKey();
        }
    }
}
