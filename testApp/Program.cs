using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
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
            List<string> docs = new List<string>();
            docs.Add("D:\\test\\План-график Козуб.docx");
            docs.Add("D:\\test\\План-график Коз.doc");
            docs.Add("D:\\test\\План-график Поляков.doc");
            docs.Add("D:\\test\\План-график Козуб.doc");
            docs.Add("D:\\test\\Plan-grafik_na_5_semestr.doc");
            docs.Add("D:\\test\\Not plan.docx");
            Model newModel = new Model(String.Empty, "XML");
            foreach (var doc in docs)
            {
                try
                {
                    
                    newModel.ParseDoc(doc);
                }
                catch (Exception e)
                {
                    switch (e.Message)
                    {
                        case "NULL_DATA":
                        {
                            Console.WriteLine("Проблемы с чтением файла " + doc + " возможно неверный формат.");
                            break;
                        }
                        case "CANT_ADD":
                        {
                            Console.WriteLine("Проблемы записью файла " + doc + " возможно ошибка в xml файле.");
                            break;
                        }
                        default:
                        {
                            Console.WriteLine(e);
                            break;
                        }
                    }
                    
                }
            }
                     
            stopwatch.Stop();
            Console.WriteLine("End. Total time: " + stopwatch.Elapsed);

            List<DataUnit> list = newModel.GetList();
            foreach (var dataUnit in list)
            {
                Console.WriteLine(dataUnit.Owner);
                var dict = dataUnit.PlanDictionary.OrderBy(pair => pair.Value).ToDictionary(pair => pair.Key, pair => pair.Value);
                foreach (var point in dict)
                {
                    Console.WriteLine(point.Key + " " + point.Value.ToString("dd.MM.yyyy", CultureInfo.InvariantCulture));
                }
            }
            
            Console.ReadKey();
        }
    }
}
