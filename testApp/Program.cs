using System;
using System.Collections.Generic;
using System.Diagnostics;
using AnalisWordLib;

namespace testApp
{
    class Program
    {
        public static void PrintInConsole(string str, int curr, int max)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            if (str.ToCharArray()[0] == 'W')
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
            }
            if (str.ToCharArray()[0] == 'E')
            {
                Console.ForegroundColor = ConsoleColor.Red;
            }
            
            Console.WriteLine(curr + "\\" + max + " " + str);
            Console.ResetColor();
        }

        public static void PrintReport(string str)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            if (str.ToCharArray()[0] == 'W')
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
            }
            if (str.ToCharArray()[0] == 'E')
            {
                Console.ForegroundColor = ConsoleColor.Red;
            }

            Console.WriteLine(str);
            Console.ResetColor();
        }

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
                        if (dirr == "")
                        {
                            dirr = "D:\\test\\";
                        }
                        stopwatch.Start();
                        newModel.ParseDir(dirr, PrintInConsole);
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
                        if (doc == "")
                        {
                            doc = "D:\\test\\План-график Козуб.docx";
                        }
                        stopwatch.Start();
                        PrintReport(newModel.ParseDoc(doc));
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
                        Console.ForegroundColor = ConsoleColor.White;
                        foreach (var stud in studList)
                        {
                            Console.WriteLine(stud);   
                        }
                        Console.ResetColor();
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
                            Console.ForegroundColor = ConsoleColor.White;
                            Console.WriteLine(student.ToString());
                            Console.ResetColor();
                        }
                        else
                        {
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine("Not found");
                            Console.ResetColor();
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
                        PrintReport(newModel.DeleteStudent(name));
                        stopwatch.Stop();
                        Console.WriteLine("Time: " + stopwatch.Elapsed);
                        Console.ReadKey();
                        Console.Clear();
                        continue;
                    }
                    case "0":
                    {
                        stopwatch = new Stopwatch();
                        Console.WriteLine("Are you sure? (y,n)");
                        string qLine = Console.ReadLine();
                        switch (qLine)
                        {
                            case "y":
                            {
                                stopwatch.Start();
                                PrintReport(newModel.DeleteAll());
                                stopwatch.Stop();
                                Console.WriteLine("Time: " + stopwatch.Elapsed);
                                Console.ReadKey();
                                break;
                            }
                            case "n":
                            {
                                Console.WriteLine("Okay.");
                                Console.ReadKey();
                                break;
                            }
                            default:
                            {
                                Console.WriteLine("Wrong input.");
                                Console.ReadKey();
                                break;
                            }
                        }
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
                        PrintReport(newModel.SetCheck(name, point, date));
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
                        PrintReport(newModel.DeleteCheck(name, point));
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
