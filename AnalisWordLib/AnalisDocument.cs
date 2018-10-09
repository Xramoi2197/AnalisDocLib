using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace AnalisWordLib
{
    public sealed class DataUnit //Класс описывающий формат данных план-графика
    {
        public string Dir { set; get; } // Файл, из которого был загружен план-график
        public Dictionary<string, DateTime> PlanDictionary { set; get; } // План-график в словаре: ключ - название точки, значение - дата
        public Dictionary<string, DateTime> CheckDictionary { set; get; }
        public string Owner { set; get; } // Фамилия И.О. слушателя

        public DataUnit() // Конструктор без параметров, выделение памяти под поля
        {
            Dir = string.Empty;
            PlanDictionary = new Dictionary<string, DateTime>();
            CheckDictionary = new Dictionary<string, DateTime>();
            Owner = string.Empty;
        }

        public bool IsNull() // Если поле Owner пустое то true
        {
            if (Owner == string.Empty) return true;
            return false;
        }

        public bool AddToDictionary(string name, string date, Dictionary<string, DateTime> dictionary) // Добавление контрольной точки в словарь с преобразованием строки в дату
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

            if (date.Length == 7)
            {
                date = "0" + date;
                pattern = "dd.MM.yy";
            }

            bool rez = DateTime.TryParseExact(date, pattern, new CultureInfo("en-US"), DateTimeStyles.None, out var dateTime); // если false, то формат даты неверный
            if (rez)
            {
                dictionary.Add(name, dateTime);
            }

            return rez;
        }

        public override string ToString()
        {
            string rezult = Owner + "\r\n";
            var dict = PlanDictionary.OrderBy(pair => pair.Value).ToDictionary(pair => pair.Key, pair => pair.Value);
            int count = dict.Count;
            foreach (var point in dict)
            {
                rezult += point.Key + " " + point.Value.ToString("dd.MM.yyyy", CultureInfo.InvariantCulture);
                if (CheckDictionary.ContainsKey(point.Key))
                {
                    rezult += " " + CheckDictionary[point.Key].ToString("dd.MM.yyyy", CultureInfo.InvariantCulture);
                }
                
                if (count != 1)
                {
                    rezult += "\r\n";
                }

                count--;
            }
            return rezult;
        }
    }

    public abstract class DataStorage<T> // Абстрактный класс для описания хранилища данных xml/DB
    {
        public abstract string TryAdd(T data);
        public abstract List<T> GetDataList();
        public abstract T FindByName(string name);
        public abstract List<string> GetStudents();
        public abstract bool DeleteData(string name);
        public abstract void DeleteAll();
        public abstract int Count();
        public abstract bool SetCheck(string name, string key, string date);
        public abstract bool DeleteCheck(string name, string key);
    }

    public sealed class XmlStorage : DataStorage<DataUnit>
    {
        private readonly XDocument _storage; // Ссылка на xml файл
        private readonly XElement _students; // Ссылка на корневой узел
        private readonly string _storageName; // Имя xml файла

        public XmlStorage(string storageName) // Открываем xml файл с именем storageName и находим корневой элемент
        {
            try
            {
                _storageName = storageName;
                _storage = XDocument.Load(_storageName);
                _students = _storage.Element("students");
            }
            catch (Exception e) // Ловим исключение, если файл не открыт
            {
                throw new Exception(e.Message, e.InnerException);
            }           
        }

        public override string TryAdd(DataUnit data) // Добавляем один план-график в xml файл
        {
            if (_students == null) // если не найден корневой элемент, то работа будет некорректна
            {
                return null;
            }

            if (TryFind(data.Owner)) // Проверяем на повтор слушателя в xml
            {
                return "Данный слушатель уже существует: " + data.Owner + " файл: " + data.Dir;
            }
            // Записываем данные в xml
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
            _storage.Save(_storageName); // Сохраняем изменения
            return "Слушатель " + data.Owner + " успешно добавлен. Файл: " + data.Dir;
        }

        public override List<DataUnit> GetDataList() // Считываем весь список из xml
        {
            List<DataUnit> rezList = new List<DataUnit>();
            foreach (var student in _students.Elements("student"))
            {
                DataUnit data = new DataUnit {Owner = student.Attribute("name")?.Value};
                // ReSharper disable once PossibleNullReferenceException
                foreach (XElement point in student.Element("points")?.Elements("point"))
                {
                    string key = point.Attribute("name")?.Value;
                    if (key == null)
                    {
                        continue;
                    }
                    string value = point.Attribute("date")?.Value;
                    data.AddToDictionary(key, value, data.PlanDictionary);
                    string check = point.Attribute("check")?.Value;
                    if (check != null)
                    {
                        data.AddToDictionary(key, check, data.CheckDictionary);
                    }

                }
                rezList.Add(data);
            }

            return rezList;
        }

        public override DataUnit FindByName(string name) // Функция поиска слушателя по имени (необходимо реализовать поиск по сходству со строкой)
        {
            DataUnit data = new DataUnit();
            foreach (var student in _students.Elements("student"))
            {
                if (name != student.Attribute("name")?.Value)
                {
                    continue;
                }
                data.Owner = student.Attribute("name")?.Value;
                // ReSharper disable once PossibleNullReferenceException
                foreach (XElement point in student.Element("points")?.Elements("point"))
                {
                    string key = point.Attribute("name")?.Value;
                    if (key == null)
                    {
                        continue;
                    }
                    string value = point.Attribute("date")?.Value;
                    data.AddToDictionary(key, value, data.PlanDictionary);
                    string check = point.Attribute("check")?.Value;
                    if (check != null)
                    {
                        data.AddToDictionary(key, check, data.CheckDictionary);
                    }
                }
                return data;
            }

            return null;
        }

        public override List<string> GetStudents()
        {
            List<string> rezList = new List<string>();
            foreach (var student in _students.Elements("student"))
            {
                string name = student.Attribute("name")?.Value;               
                rezList.Add(name);
            }

            return rezList;
        }

        public override bool DeleteData(string name) // Удаление одного план-графика
        {
            XElement deleted = null;
            foreach (var student in _students.Elements("student"))
            {
                if (name != student.Attribute("name")?.Value)
                {
                    continue;
                }

                deleted = student;
                break;
            }

            if (deleted == null)
            {
                return false;
            }
            deleted.Remove();
            _storage.Save(_storageName);
            return true;
        }

        public override void DeleteAll()
        {
            _students.RemoveAll();
            _storage.Save(_storageName);
        }

        public override int Count()
        {
            var count = _students.Elements("student").Count();
            return count;
        }

        public override bool SetCheck(string name, string key, string date)
        {
            DataUnit data = FindByName(name);
            string pattern = "dd.MM.yyyy";

            if (date.Length == 9)
            {
                date = "0" + date;
            }

            if (date.Length == 8)
            {
                pattern = "dd.MM.yy";
            }

            if (date.Length == 7)
            {
                date = "0" + date;
                pattern = "dd.MM.yy";
            }

            if (!data.PlanDictionary.ContainsKey(key))
            {
                return false;
            }
            bool rez = DateTime.TryParseExact(date, pattern, new CultureInfo("en-US"), DateTimeStyles.None, out _);
            if (!rez)
            {
                return false;
            }
            foreach (var student in _students.Elements("student"))
            {
                if (name != student.Attribute("name")?.Value)
                {
                    continue;
                }
                // ReSharper disable once PossibleNullReferenceException
                foreach (XElement point in student.Element("points")?.Elements("point"))
                {
                    string pointName = point.Attribute("name")?.Value;
                    if (pointName != key)
                    {
                        continue;
                    }

                    if (point.Attribute("check")?.Value != null)
                    {
                        // ReSharper disable once PossibleNullReferenceException
                        point.Attribute("check").Value = date;
                        _storage.Save(_storageName);
                        return true;
                    }
                    point.Add(new XAttribute("check", date));
                    _storage.Save(_storageName);
                    return true;
                }
            }
            return false;
        }

        public override bool DeleteCheck(string name, string key)
        {
            DataUnit data = FindByName(name);
            if (!data.CheckDictionary.ContainsKey(key))
            {
                return false;
            }
            foreach (var student in _students.Elements("student"))
            {
                if (name != student.Attribute("name")?.Value)
                {
                    continue;
                }
                // ReSharper disable once PossibleNullReferenceException
                foreach (XElement point in student.Element("points")?.Elements("point"))
                {
                    string pointName = point.Attribute("name")?.Value;
                    if (pointName != key)
                    {
                        continue;
                    }

                    point.Attribute("check")?.Remove();
                    _storage.Save(_storageName);
                    return true;
                }
            }
            return false;
        }

        public bool TryFind(string name) // Проверка существования слушателя в xml
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
        private static ConfClass _config; // Класс конфигурации приложения
        //private static AnalisDocument<DataUnit> _analis; // Класc метода анализа план-графиков
        private static DataStorage<DataUnit> _storage; // Класс для хранения данных

        public Model(string header, string storageType, string defDirr) // Задаем модель header - заголовок план-графика, storageType - XML/DB
        {
            try
            {
                _config = new Configuration(header, storageType, defDirr);
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

        public List<string> ParseDir(string dirrectory)
        {
            string[] documents = Directory.GetFiles(_config.GetDefDirr(), "*.doc?");
            List<string> rez = new List<string>();
            foreach (string doc in documents)
            {
                rez.Add(ParseDoc(doc));
            }

            return rez;
        }

        public string ParseDoc(string document) // Анализ одного документа
        {
            AnalisDocument<DataUnit> analis;
            string rez = null;
            var reg = new Regex(@".\w+$");
            var match = reg.Match(document);
            var extension = match.Value;
            if (extension == ".doc" || extension == ".docx") // Выбор метода в зависимости от расширения
            {
                analis = new AnalisDoc(_config.GetHeader());
                DataUnit data = analis.Parse(document);
                
                if (data.IsNull()) // Исключение, если метод для анализа вернул пустой объект
                {
                    rez = "Проблемы с чтением файла " + document + " возможно неверный формат.";
                    return  rez;
                }

                rez = _storage.TryAdd(data);
                
            }
            if (rez == null) //Если не удалось поместить в хранилище
            {
                rez = "Проблемы записью файла " + document + " возможно ошибка в xml файле.";
            }
            return rez;
        }

        public List<DataUnit> GetList()
        {
            var dataList = _storage.GetDataList();
            return dataList;
        }

        public List<string> GetStudentsFromStorage()
        {
            var students = _storage.GetStudents();
            return students;
        }

        public DataUnit FindStudent(string name)
        {
            var student = _storage.FindByName(name);
            return student;
        }

        public string DeleteStudent(string name)
        {
            string rez;
            if (!_storage.DeleteData(name))
            {
                rez = "Данный слушатель " + name + " не найден.";
            }
            else
            {
                rez = "План-график слушателя " + name + " удален.";
            }

            return rez;
        }

        public string DeleteAll()
        {
            string rez;
            if (_storage.Count() > 0)
            {
                _storage.DeleteAll();
                rez = "Хранилище очищено.";
                return rez;
            }

            rez = "Хранилище пусто";
            return rez;
        }

        public string SetCheck(string name, string point, string date)
        {
            string rez;

            if (!_storage.SetCheck(name, point, date))
            {
                rez = "Не удалось добавить отметку у " + name + " " + point + " " + date;
            }
            else
            {
                rez = "Слушатель " + name + " сдал " + point + " " + date;
            }

            return rez;
        }

        public string DeleteCheck(string name, string point)
        {
            string rez;
            if (!_storage.DeleteCheck(name, point))
            {
                rez = "Не удалось удалить отметку о проверке у " + name + " в " + point;
            }
            else
            {
                rez = "Отметка о " + name + " о " + point + " удалена.";
            }

            return rez;
        }
    }

    public abstract class ConfClass // Абстрактный класс описываюший работу конфигураций приложения
    {
        public abstract void SetHeader(string header);
        public abstract string GetHeader();
        public abstract void SetStorageType(string storageType);
        public abstract string GetStorageType();
        public abstract void SetDefDirr(string dirrectory);
        public abstract string GetDefDirr();

    }

    public sealed class Configuration : ConfClass
    {
        private string _header;
        private string _storageType;
        private string _defaultDirrectory;

        public Configuration(string header, string storageType, string defaultDirrectory)
        {
            _header = header;
            if (storageType != "XML" && storageType != "DB")
            {
                throw new Exception("WRONG_STORAGE_TYPE");
            }
            _storageType = storageType;
            _defaultDirrectory = defaultDirrectory;
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

        public override void SetDefDirr(string dirrectory)
        {
            _defaultDirrectory = dirrectory;
        }

        public override string GetDefDirr()
        {
            return _defaultDirrectory;
        }
    } // Класс с параметрами работы приложения

    public abstract class AnalisDocument<T> // Абстрактный класс объединяющий алгоритмы анализа план-графиков
    {
        public abstract T Parse(string document);
    }

    public sealed class AnalisDoc : AnalisDocument<DataUnit>
    {
        private static Word.Application _wordApp; // Объект Application
        private static Word.Document _wordDocument; // Объект Document
        private static DataUnit _data; // Объект данных с результатом
        private static string _header; // Заголовок план-графика

        public AnalisDoc(string header) // Коструктор инициализирует объект Application и заполняет заголовок для поиска
        {
            _wordApp = new Word.Application { Visible = false }; // Приложение открывается в фоне
            _wordDocument = new Word.Document();
            _header = header == string.Empty ? "ПЛАН-ГРАФИК" : header; // Заголовок по умолчанию "ПЛАН-ГРАФИК"

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
                //Закрываем объект Application
                CloseWord();               
            }
            return _data;
        }

        private static void OpenWord(object fileName) // Метод открывает конкретный файл MS Word
        {
            try
            {
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

        private static void CloseWord() // Метод закрывает файл
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

        private static void AnaliticWord() // Метод перебирает файл и вызывает метод анализа
        {
            var counter = 0; // Счетчик для выбора условий проверки
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
                if (!_data.AddToDictionary(jobs.ToArray()[i], dates.ToArray()[i], _data.PlanDictionary))
                {
                    _data.Owner = string.Empty;
                    throw new Exception();
                }
                

            }
        }

        private static int CheckText(Word.Range text, ref int counter, ref Stack<string> jobs, ref Stack<string> dates) // Метод анализа
        {
            // Получаем номер сраницы объекта
            object vobjN = text.Information[Word.WdInformation.wdActiveEndPageNumber];
            var strNum = vobjN.ToString();
            int num = int.Parse(strNum);

            switch (counter)
            {
                case 0: // Поиск заголовка
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
                            counter++; // Если заголовок найден, то переходим к следующему условию
                            return 0;
                        }

                        return -1;
                    }
                case 1: // Поиск в таблице
                    {
                        if (text.Tables.Count > 0) // Находим таблицу
                        {
                            // Получаем номер ячейки
                            object stableInfo = text.Information[Word.WdInformation.wdStartOfRangeRowNumber];
                            object ttableInfo = text.Information[Word.WdInformation.wdStartOfRangeColumnNumber];
                            var s = text.Text;
                            if ((int)stableInfo == 1 && (int)ttableInfo == 1 && num > 1) // Если таблица на 2 и более странице, то не соответствие формату
                            {
                                return -2;
                            }
                            if ((int)stableInfo > 1 && (int)ttableInfo < 4 && (int)ttableInfo > 1) // Парсим таблицу
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

                        if (text.Tables.Count == 0 && jobs.Count > 0) // Если таблица закончилась, переходим к следующему заданию
                        {
                            counter++;
                            return 0;
                        }

                        return -1;
                    }
                case 2: // Посик ФИО
                    {
                        var reg = new Regex(@"\w\.\w\.\s+\w+\b");
                        var match = reg.Match(text.Text);
                        if (match.Success) // Находим И.О. Фамилия и преобразуем к Фамилия И.О.
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

