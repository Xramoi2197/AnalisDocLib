using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace AnalisWordLib
{
    public sealed class DataUnit //Класс описывающий формат данных план-графика
    {
        public string Dir { set; get; }
        public Dictionary<string, DateTime> PlanDictionary { set; get; }
        public string Owner { set; get; }

        public DataUnit()
        {
            Dir = string.Empty;
            PlanDictionary = new Dictionary<string, DateTime>();
            Owner = string.Empty;
        }

        public bool IsNull()
        {
            if (Owner == string.Empty) return true;
            return false;
        }

        public static bool IsValidName(string name)
        {
            const string namePattern = @"^[А-Я][a-я]+\s[А-Я]\.[А-Я]\.$";
            var reg = new Regex(namePattern);
            var match = reg.Match(name);
            if (!match.Success)
            {
                return false;
            }
            return true;
        }

        public bool AddToPlan(string name, string date)
        {
            string pattern = "dd.MM.yyyy";
            if (date.Length == 9)
            {
                date = "0" + date;
            }

            if (date.Length == 8)
            {
                pattern = "dd.MM.yy";
            }
            DateTime dateTime = new DateTime();
            CultureInfo enUs = new CultureInfo("en-US");
            bool rez = DateTime.TryParseExact(date, pattern, new CultureInfo("en-US"), DateTimeStyles.None, out dateTime);
            //dateTime = DateTime.TryParseExact(date, "dd.MM.yyyy", null);
            if (rez)
            {
                PlanDictionary.Add(name, dateTime);
            }

            return rez;
        }
    }

    public abstract class DataStorage<T> // Абстрактный класс для описания хранилища данных xml/DB
    {
        public abstract string TryAdd(T data);
        public abstract List<T> GetDataList();
        public abstract T FindByName(string name);
    }

    public sealed class XmlStorage : DataStorage<DataUnit>
    {
        private readonly XDocument _storage;
        private readonly XElement _students;
        private readonly string _storageName;

        public XmlStorage(string storageName)
        {
            try
            {
                _storageName = storageName;
                _storage = XDocument.Load(_storageName);
                _students = _storage.Element("students");
            }
            catch (Exception e)
            {
                throw new Exception(e.Message, e.InnerException);
            }           
        }

        public override string TryAdd(DataUnit data)
        {
            if (_students == null)
            {
                return null;
            }

            if (TryFind(data.Owner))
            {
                return "Данный слушатель уже существует: " + data.Owner + " файл: " + data.Dir;
            }

            _students.Add(new XElement("student",
                new XAttribute("name", data.Owner),
                new XAttribute("dirr", data.Dir),
                new XElement("points")));
            XElement points = null;
            foreach (var student in _students.Elements("student"))
            {
                if (student.Attribute("name")?.Value == data.Owner)
                {
                    points = student.Element("points");
                }
            }

            foreach (var plan in data.PlanDictionary)
            {
                points?.Add(new XElement("point",
                    new XAttribute("name", plan.Key),
                    new XAttribute("date", plan.Value.Date.ToString("dd.MM.yyyy", CultureInfo.InvariantCulture))));
            }
            _storage.Save(_storageName);
            return "Слушатель " + data.Owner + " успешно добавлен. Файл: " + data.Dir;
        }

        public override List<DataUnit> GetDataList()
        {
            List<DataUnit> rezList = new List<DataUnit>();
            foreach (var student in _students.Elements("student"))
            {
                DataUnit data = new DataUnit();
                data.Owner = student.Attribute("name")?.Value;
                foreach (var point in student.Element("points").Elements("point"))
                {
                    string key = point.Attribute("name")?.Value;
                    if (key == null)
                    {
                        continue;
                    }
                    string value = point.Attribute("date")?.Value;
                    try
                    {
                        data.AddToPlan(key, value);
                    }
                    catch (Exception e)
                    {
                        throw new Exception(e.Message,e.InnerException);
                    }
                    
                }
                rezList.Add(data);
            }

            return rezList;
        }

        public override DataUnit FindByName(string name)
        {
            return null;
        }

        public bool TryFind(string name)
        {
            foreach (var student in _students.Elements("student"))
            {
                if (student.Attribute("name")?.Value == name)
                {
                    return true;
                }
            }
            return false;
        }
    } // Класс для хранения данных в xml

    public sealed class Model // Фасад модуля описывает логику взаимодействя классов и предоставляет методы для внешних модулей
    {
        private static ConfClass _config;
        private static AnalisDocument<DataUnit> _analis;
        private static DataStorage<DataUnit> _storage;

        public Model(string header, string storageType)
        {
            try
            {
                _config = new Configuration(header, storageType);
                if (_config.GetStorageType() == "XML")
                {
                    _storage = new XmlStorage("ModelStorage.xml");
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
            if (extension == ".doc" || extension == ".docx")
            {
                _analis = new AnalisDoc(_config.GetHeader());
                DataUnit data = _analis.Parse(document);
                
                if (data.IsNull())
                {
                    throw new Exception("NULL_DATA");
                }

                string rez = _storage.TryAdd(data);
                if (rez == null)
                {
                    throw new Exception("CANT_ADD");
                }
                Console.WriteLine(rez);

            }
        }

        public List<DataUnit> GetList()
        {
            var dataList = _storage.GetDataList();
            return dataList;
        }
    }

    public abstract class ConfClass // Абстрактный класс описываюший работу конфигураций приложения
    {
        public abstract void SetHeader(string header);
        public abstract string GetHeader();
        public abstract void SetStorageType(string storageType);
        public abstract string GetStorageType();

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

        public override void SetStorageType(string storageType)
        {
            if (storageType != "XML" && storageType != "DB")
            {
                throw new Exception("WRONG_STORAGE_TYPE");
            }
            _storageType = storageType;
        }

        public override string GetStorageType()
        {
            return _storageType;
        }
    } // Класс с параметрами работы приложения

    public abstract class AnalisDocument<T> // Абстрактный класс объединяющий алгоритмы анализа план-графиков
    {
        public abstract T Parse(string document);
    }

    public sealed class AnalisDoc : AnalisDocument<DataUnit>
    {
        private static Word.Application _wordApp;
        private static Word.Document _wordDocument;
        private static DataUnit _data;
        private static string _header;

        public AnalisDoc(string header)
        {
            _wordApp = new Word.Application { Visible = false };
            _wordDocument = new Word.Document();
            _header = header == string.Empty ? "ПЛАН-ГРАФИК" : header;

            _data = new DataUnit();
        }

        public override DataUnit Parse(string document)
        {
            try
            {
                _data.Dir = document;
                //Открываем ворд файл
                OpenWord(document);
                //Извлекаем нужные поля
                AnaliticWord();           
            }
            catch (Exception)
            {
                // ignored
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
                throw new Exception(ex.Message, ex.InnerException);
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
            var counter = 0;
            Stack<string> jobs = new Stack<string>();
            Stack<string> dates = new Stack<string>();
            for (var i = 1; i < _wordDocument.Paragraphs.Count; i++)
            {
                var s = _wordDocument.Paragraphs[i].Range;

                var rez = CheckText(s, ref counter, ref jobs, ref dates);

                if (rez < -1)
                {
                    return;
                }
            }

            for (int i = 0; i < jobs.Count; i++)
            {
                if (!_data.AddToPlan(jobs.ToArray()[i], dates.ToArray()[i]))
                {
                    _data.Owner = string.Empty;
                    throw new Exception();
                }
                

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

