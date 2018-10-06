﻿using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;

namespace AnalisWordLib
{


    public sealed class DataUnit //Класс описывающий формат данных план-графика
    {
        public string Dir { set; get; }
        public Dictionary<string, string> PlanDictionary { set; get; }
        public string Owner { set; get; }

        public DataUnit()
        {
            Dir = string.Empty;
            PlanDictionary = new Dictionary<string, string>();
            Owner = string.Empty;
        }

        public bool IsNull()
        {
            if (Owner == string.Empty) return true;
            return false;
        }
    }


    public abstract class DataStorage // Абстрактный класс для описания хранилища данных xml/DB
    {
        public abstract string TryAdd(DataUnit data);
        
    }

    public sealed class XmlStorage : DataStorage
    {
        public override string TryAdd(DataUnit data)
        {
            throw new NotImplementedException();
        }

        private bool CheckData(DataUnit data)
        {
            return true;
        }
    } // Класс для хранения данных в xml

    public sealed class Model // Фасад модуля описывает логику взаимодействя классов и предоставляет методы для внешних модулей
    {
        private static ConfClass _config;
        private static AnalisDocument _analis;
        private static DataStorage _storage;
        private const string DataError = "NULL_DATA";

        public Model(string header, string storageType)
        {
            try
            {
                _config = new Configuration(header, storageType);
                if (_config.GetStorage() == "XML")
                {
                    _storage = new XmlStorage();
                }
            }
            catch (Exception e)
            {
                throw new Exception(e.Message, e.InnerException);
            }               
        }

        public void ParseDoc(string document)
        {
            var reg = new Regex(@".\w+$");
            var match = reg.Match(document);
            var extension = match.Value;
            if (extension == ".doc")
            {
                _analis = new AnalisDoc(_config.GetHeader());
                try
                {
                    DataUnit data = _analis.Parse(document);
                    if (data.IsNull())
                    {
                        throw new Exception(DataError);
                    }
                }
                catch (Exception e)
                {
                    throw new Exception(e.Message, e.InnerException);
                }
                
                
            }
        }
    }

    public abstract class ConfClass // Абстрактный класс описываюший работу конфигураций приложения
    {
        public abstract void SetHeader(string header);
        public abstract string GetHeader();
        public abstract void SetStorage(string storageType);
        public abstract string GetStorage();

    }

    public sealed class Configuration : ConfClass
    {
        private static string _header;
        private static string _storageType;

        public Configuration(string header, string storageType)
        {
            _header = header;
            if (storageType != "XML" && storageType != "DB")
            {
                throw new Exception("WRONG_STORAGE_TYPE");
            }
            _storageType = storageType;
        }

        public override void SetHeader(string header)
        {
            _header = header;
        }

        public override string GetHeader()
        {
            return _header;
        }

        public override void SetStorage(string storageType)
        {
            if (storageType != "XML" && storageType != "DB")
            {
                throw new Exception("WRONG_STORAGE_TYPE");
            }
            _storageType = storageType;
        }

        public override string GetStorage()
        {
            return _storageType;
        }
    } // Класс с параметрами работы приложения

    public abstract class AnalisDocument // Абстрактный класс объединяющий алгоритмы анализа план-графиков
    {
        public abstract DataUnit Parse(string document);
    }

    public sealed class AnalisDoc : AnalisDocument
    {
        private static Word.Application _wordApp;
        private static Word.Document _wordDocument;
        private static DataUnit _data;
        private static string _header;
        public const string FormatError = "WRONG_FORMAT";
        public const string AnalisError = "CANT_ANALIS";
        public const string OpenError = "CANT_OPEN";

        public AnalisDoc(string header)
        {
            _wordApp = new Word.Application { Visible = false };
            _wordDocument = new Word.Document();
            if (header == string.Empty)
            {
                _header = "ПЛАН-ГРАФИК";
            }
            else _header = header;

            _data = new DataUnit();
        }

        public override DataUnit Parse(string document)
        {
            try
            {
                //Открываем ворд файл
                OpenWord(document);
                //Извлекаем нужные поля
                AnaliticWord();
            }
            catch (Exception e)
            {
                switch (e.Message)
                {
                    case OpenError:
                    {
                        throw new Exception(OpenError);
                    }
                    case AnalisError:
                    {
                        throw new Exception(AnalisError);
                    }
                    case FormatError:
                    {
                        throw new Exception(FormatError);
                    }
                    default:
                    {
                        throw new Exception(e.Message, e.InnerException);
                    }
                }
            }
            finally
            {
                CloseWord();
                
            }
            return _data;
        }

        private static void OpenWord(object fileName)
        {
            try
            {

                //не отображаем
                //Прописываем свойства открытия

                object confirmConversions = true;
                object readOnly = true;
                object addToRecentFiles = true;
                var passwordDocument = Type.Missing;
                var passwordTemplate = Type.Missing;
                object revert = false;
                var writePasswordDocument = Type.Missing;
                var writePasswordTemplate = Type.Missing;
                var format = Type.Missing;
                var encoding = Type.Missing;
                var oVisible = Type.Missing;
                var openAndRepair = Type.Missing;
                var documentDirection = Type.Missing;
                object noEncodingDialog = false;
                var xmlTransform = Type.Missing;
                //Открываем
                _wordDocument = _wordApp.Documents.Open(ref fileName,
                    ref confirmConversions, ref readOnly, ref addToRecentFiles,
                    ref passwordDocument, ref passwordTemplate, ref revert,
                    ref writePasswordDocument, ref writePasswordTemplate,
                    ref format, ref encoding, ref oVisible,
                    ref openAndRepair, ref documentDirection, ref noEncodingDialog, ref xmlTransform);
            }
            catch (Exception ex)
            {
                throw new Exception(OpenError, ex.InnerException);
            }
        }

        private static void CloseWord()
        {
            if (_wordDocument != null)
            {
                try
                {
                    _wordDocument.Close(false);
                }
                finally
                {
                    _wordDocument = null;
                }
            }
            if (_wordApp != null)
            {
                try
                {
                    _wordApp.Quit(false);
                }
                finally
                {
                    _wordApp = null;
                }

            }
        }

        private static void AnaliticWord()
        {
            try
            {
                var counter = 0;
                Stack<string> jobs = new Stack<string>();
                Stack<string> dates = new Stack<string>();
                for (var i = 1; i < _wordDocument.Paragraphs.Count; i++)
                {
                    var s = _wordDocument.Paragraphs[i].Range;

                    var rez = CheckText(s, ref counter, ref jobs, ref dates);

                    if (rez < -1)
                    {
                        throw new Exception(FormatError);
                    }
                }

                for (int i = 0; i < jobs.Count; i++)
                {
                    _data.PlanDictionary.Add(jobs.ToArray()[i], dates.ToArray()[i]);
                }
            }
            catch (Exception ex)
            {
                throw new Exception(AnalisError, ex.InnerException);
            }
        }

        private static int CheckText(Word.Range text, ref int counter, ref Stack<string> jobs, ref Stack<string> dates)
        {
            object vobjN = text.Information[Word.WdInformation.wdActiveEndPageNumber];
            var strNum = vobjN.ToString();
            int num = int.Parse(strNum);

            switch (counter)
            {
                case 0:
                    {
                        if (num > 1)
                        {
                            return -2;
                        }

                        string pattern = @"^" + _header;
                        var reg = new Regex(pattern);
                        var match = reg.Match(text.Text);
                        if (match.Success)
                        {
                            counter++;
                            return 0;
                        }

                        return -1;
                    }
                case 1:
                    {
                        if (text.Tables.Count > 0)
                        {
                            object stableInfo = text.Information[Word.WdInformation.wdStartOfRangeRowNumber];
                            object ttableInfo = text.Information[Word.WdInformation.wdStartOfRangeColumnNumber];
                            var s = text.Text;
                            if ((int)stableInfo == 1 && (int)ttableInfo == 1 && num > 1)
                            {
                                return -2;
                            }
                            if ((int)stableInfo > 1 && (int)ttableInfo < 4 && (int)ttableInfo > 1)
                            {
                                if ((int)ttableInfo == 2)
                                {
                                    jobs.Push(s.Remove(s.Length - 2, 2));
                                }
                                else
                                {
                                    dates.Push(s.Remove(s.Length - 2, 2));
                                }
                            }
                        }

                        if (text.Tables.Count == 0 && jobs.Count > 0)
                        {
                            counter++;
                            return 0;
                        }

                        return -1;
                    }
                case 2:
                    {
                        var reg = new Regex(@"\w\.\w\.\s+\w+\b");
                        var match = reg.Match(text.Text);
                        if (match.Success)
                        {
                            counter++;
                            string name = match.Value;
                            reg = new Regex(@"\w+$");
                            match = reg.Match(name);
                            _data.Owner = match.Value + " ";
                            reg = new Regex(@"\w\.\w\.");
                            match = reg.Match(name);
                            _data.Owner += match.Value;
                            return 0;
                        }

                        return -1;
                    }
            }

            return 0;
        }
    } // Класс для анализа план-графиков с расширением .doc и возможностью анализа .docx
}
