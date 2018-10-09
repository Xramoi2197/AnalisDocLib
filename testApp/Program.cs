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
            Model newModel = new Model(string.Empty, "XML");
            while (true)
            {
                Console.WriteLine("Menu:");
                Console.WriteLine("1.Add test to xml");
                Console.WriteLine("2.Print data from xml");
                Console.WriteLine("3.Get students");
                Console.WriteLine("4.Find by name");
                Console.WriteLine("5.Delete by name");
                Console.WriteLine("6.Delete all");
                Console.WriteLine("7.Set check");
                Console.WriteLine("8.Remove check");
                Console.WriteLine("enter - Exit");
                Console.Write("Input:");
                var userInput = Console.ReadLine();
                Stopwatch stopwatch;
                switch (userInput)
                {
                    case "1":
                    {
                        stopwatch = new Stopwatch();
                        Console.WriteLine("Start.");
                        stopwatch.Start();
                        List<string> docs = new List<string>();
                        docs.Add("D:\\test\\План-график Козуб.docx");
                        docs.Add("D:\\test\\План-график Коз.doc");
                        docs.Add("D:\\test\\План-график Поляков.doc");
                        docs.Add("D:\\test\\План-график Козуб.doc");
                        docs.Add("D:\\test\\Plan-grafik_na_5_semestr.doc");
                        docs.Add("D:\\test\\Not plan.docx");
                        
                        foreach (var doc in docs)
                        {
                            string rez = newModel.ParseDoc(doc);
                            Console.WriteLine(rez);
                        }

                        stopwatch.Stop();
                        Console.WriteLine("End. Total time: " + stopwatch.Elapsed);
                        Console.ReadKey();
                        Console.Clear();
                            continue;
                    }
                    case "2":
                    {
                        stopwatch = new Stopwatch();
                        stopwatch.Start();
                        List<DataUnit> list = newModel.GetList();
                        foreach (var dataUnit in list)
                        {
                            Console.WriteLine(dataUnit.ToString());
                        }
                        stopwatch.Stop();
                        Console.WriteLine("Time: " + stopwatch.Elapsed);
                        Console.ReadKey();
                        Console.Clear();
                            continue;
                    }
                    case "3":
                    {
                        stopwatch = new Stopwatch();
                        stopwatch.Start();
                        List<string> studList = newModel.GetStudentsFromStorage();
                        foreach (var stud in studList)
                        {
                            Console.WriteLine(stud);   
                        }
                        stopwatch.Stop();
                        Console.WriteLine("Time: " + stopwatch.Elapsed);
                        Console.ReadKey();
                        Console.Clear();
                            continue;
                    }
                    case "4":
                    {
                        stopwatch = new Stopwatch();
                        Console.Write("Add name: ");
                        var name = Console.ReadLine();
                        stopwatch.Start();
                        var student = newModel.FindStudent(name);
                        if (student != null)
                        {
                            Console.WriteLine(student.ToString());
                        }
                        else
                        {
                            Console.WriteLine("Not found");
                        }
                        stopwatch.Stop();
                        Console.WriteLine("Time: " + stopwatch.Elapsed);
                        Console.ReadKey();
                        Console.Clear();
                            continue;
                    }
                    case "5":
                    {
                        stopwatch = new Stopwatch();
                        Console.Write("Add name: ");
                        var name = Console.ReadLine();
                        stopwatch.Start();
                        Console.WriteLine(newModel.DeleteStudent(name));
                        stopwatch.Stop();
                        Console.WriteLine("Time: " + stopwatch.Elapsed);
                        Console.ReadKey();
                        Console.Clear();
                        continue;
                    }
                    case "6":
                    {
                        stopwatch = new Stopwatch();
                        stopwatch.Start();
                        Console.WriteLine(newModel.DeleteAll());
                        stopwatch.Stop();
                        Console.WriteLine("Time: " + stopwatch.Elapsed);
                        Console.ReadKey();
                        Console.Clear();
                        continue;
                    }
                    case "7":
                    {
                        stopwatch = new Stopwatch();
                        Console.Write("Add name: ");
                        var name = Console.ReadLine();
                        Console.Write("Add point: ");
                        var point = Console.ReadLine();
                        Console.Write("Add date: ");
                        var date = Console.ReadLine();
                        stopwatch.Start();
                        Console.WriteLine(newModel.SetCheck(name, point, date));
                        stopwatch.Stop();
                        Console.WriteLine("Time: " + stopwatch.Elapsed);
                        Console.ReadKey();
                        Console.Clear();
                        continue;
                    }
                    case "8":
                    {
                        stopwatch = new Stopwatch();
                        Console.Write("Add name: ");
                        var name = Console.ReadLine();
                        Console.Write("Add point: ");
                        var point = Console.ReadLine();
                        stopwatch.Start();
                        Console.WriteLine(newModel.DeleteCheck(name, point));
                        stopwatch.Stop();
                        Console.WriteLine("Time: " + stopwatch.Elapsed);
                        Console.ReadKey();
                        Console.Clear();
                        continue;
                    }
                    case "":
                    {
                        return;
                    }
                    default:
                    {
                        Console.WriteLine("Wrong input");
                        Console.ReadKey();
                        Console.Clear();
                        continue;
                    }
                }
            }               
        }
    }
}
