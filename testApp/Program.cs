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
            Model newModel = new Model(string.Empty, "XML", "D:\\test\\");
            while (true)
            {
                Console.WriteLine("Menu:");
                Console.WriteLine("1.Add dirr to xml");
                Console.WriteLine("2.Add doc to xml");
                Console.WriteLine();
                Console.WriteLine("3.Print data from xml");
                Console.WriteLine("4.Get students");
                Console.WriteLine("5.Find by name");
                Console.WriteLine();
                Console.WriteLine("6.Set check");
                Console.WriteLine("7.Remove check");
                Console.WriteLine();
                Console.WriteLine("9.Delete by name");
                Console.WriteLine("0.Delete all");
                Console.WriteLine();
                Console.WriteLine("enter - Exit");
                Console.Write("Input:");
                var userInput = Console.ReadLine();
                Console.WriteLine();
                Stopwatch stopwatch;
                switch (userInput)
                {
                    case "1":
                    {
                        stopwatch = new Stopwatch();
                        Console.WriteLine("Start.");
                        Console.Write("Add dirrectory path: ");
                        var dirr = Console.ReadLine();
                        if (dirr == null)
                        {
                            dirr = "D:\\test\\";
                        }
                        stopwatch.Start();
                        List<string> docs = new List<string>();
                        List<string> rezList = newModel.ParseDir(dirr);
                        foreach (var rez in rezList)
                        {
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
                        Console.WriteLine("Start.");
                        Console.Write("Add document name: ");
                        var doc = Console.ReadLine();
                        if (doc == null)
                        {
                            doc = "D:\\test\\План-график Козуб.doc";
                        }
                        stopwatch.Start();
                        Console.WriteLine(newModel.ParseDoc(doc));
                        stopwatch.Stop();
                        Console.WriteLine("End. Total time: " + stopwatch.Elapsed);
                        Console.ReadKey();
                        Console.Clear();
                        continue;
                    }
                    case "3":
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
                    case "4":
                    {
                        stopwatch = new Stopwatch();
                        stopwatch.Start();
                        List<string> studList = newModel.GetStudentsFromStorage();
                        foreach (var stud in studList)
                        {
                            Console.WriteLine(stud);   
                        }
                        Console.WriteLine("Всего план-графиков записано: " + studList.Count);
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
                    case "9":
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
                    case "0":
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
                    case "6":
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
                    case "7":
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
