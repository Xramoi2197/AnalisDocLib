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
            Model newModel = new Model(String.Empty, "XML");
            while (true)
            {
                Console.WriteLine("Menu:");
                Console.WriteLine("1.Add test to xml");
                Console.WriteLine("2.Print data from xml");
                Console.WriteLine("3.Get students");
                Console.WriteLine("4.Find by name");
                Console.WriteLine("enter - Exit");
                Console.Write("Input:");
                var userInput = Console.ReadLine();
                switch (userInput)
                {
                    case "1":
                    {
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
                            try
                            {

                                string rez = newModel.ParseDoc(doc);
                                Console.WriteLine(rez);
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
                        Console.ReadKey();
                        Console.Clear();
                            continue;
                    }
                    case "2":
                    {
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
                    case "":
                    {
                        return;
                    }
                    default:
                    {
                        Console.WriteLine("Wrong input");
                        Console.Clear();
                        continue;
                    }
                }
            }               
        }
    }
}
