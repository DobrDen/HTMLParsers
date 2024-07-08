using CsvHelper;
using CsvHelper.Configuration;
using HtmlAgilityPack;
using log4net;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;


namespace HTMLParser1
{
    internal class Program
    {
      
        private static readonly ILog _log = LogManager.GetLogger(typeof(Program));
        static void Main(string[] args)
        {

            log4net.Config.XmlConfigurator.Configure();
            string url = "https://www.wienerborse.at/en/bonds/";
            int startPage = 1;
            var web = new HtmlWeb();
            //Создание списка объектов класса TableNode для хранения записей таблицы
            List<TableNode> tableNodes = new List<TableNode>();

            try
            {
                _log.Info("Начало загрузки страницы");
                //Загрузка html-страницы (1-ой страницы таблицы)
                var document = web.Load(url+$"?c7928-page={startPage}");
                //Взятие номера последней страницы таблицы
                int lastPageNum = int.Parse(document.DocumentNode.SelectSingleNode("//*[@id=\"c7928-module\"]/div/div[2]/div[1]/div[1]/div[2]/ul/li[7]").InnerText);
                _log.Info($"Номер последней страницы: {lastPageNum}");
                //Цикл перебора всех страниц таблицы
                for (int i = startPage; i <= lastPageNum; i++)
                {
                    try
                    {
                        _log.Info($"Парсинг данных страницы {i}");
                        //Получение коллекции записей из таблицы
                        //Если ни одной записи в таблице не найдено, выбрасывается исключение
                        var nodes = document.DocumentNode.SelectNodes("//*[@id=\"c7928-module-container\"]/table/tbody/tr[position()>0]") ?? throw new Exception("Не обнаружены строки в таблице");
                        //Перебор данных в каждой записи и добавление их в список объектов
                        foreach (var node in nodes)
                        {
                                tableNodes.Add(new TableNode()
                                {
                                    Name = HtmlEntity.DeEntitize(node.SelectSingleNode("td[1]").InnerText),
                                    Last = HtmlEntity.DeEntitize(node.SelectSingleNode("td[2]").InnerText),
                                    Changes = HtmlEntity.DeEntitize(node.SelectSingleNode("td[3]").InnerText),
                                    DateTime = HtmlEntity.DeEntitize(node.SelectSingleNode("td[4]").InnerText),
                                    ISIN = HtmlEntity.DeEntitize(node.SelectSingleNode("td[5]").InnerText),
                                    BidVolume = node.SelectSingleNode("td[6]").InnerHtml,
                                    AskVolume = node.SelectSingleNode("td[7]").InnerHtml,
                                    Maturity = HtmlEntity.DeEntitize(node.SelectSingleNode("td[8]").InnerText),
                                    Status = HtmlEntity.DeEntitize(node.SelectSingleNode("td[9]").InnerText)
                                });
                        }
                        url = $"https://www.wienerborse.at/en/bonds/?c7928-page={i+1}";
                        //Загрузка следующей страницы таблицы
                        document = web.Load(url);

                    }
                    catch (Exception ex)
                    {
                        _log.Error($"Ошибка загрузки или парсинга страницы {i}: {ex.Message}");
                        break; // Остановка дальнейшей загрузки, если какая-либо страница не была загружена
                    }
                }
                if (tableNodes.Count == 0)
                {
                    throw new Exception("Коллекция данных пуста, запись в CSV файл невозможна.");
                }
                _log.Info("Успешное заполнение коллекции данных");
                string fileName = $"output_{DateTime.Now:yyyy-MM-dd_HH-mm-ss}.csv";
                var config = new CsvConfiguration(CultureInfo.InvariantCulture)
                {
                    Delimiter = "\t",
                };

                try
                {
                    _log.Info("Начало записи в CSV файл");
                    //Запись полученного списка объектов класса TableNode в CSV-файл
                    using (var writer = new StreamWriter(fileName))
                    using (var csv = new CsvWriter(writer, config))
                    {
                        csv.WriteRecords(tableNodes);
                    }
                    _log.Info($"Данные успешно записаны в файл {fileName}");
                }
                catch (DirectoryNotFoundException ex)
                {
                    _log.Error($"Каталог не найден:  {ex.Message}");
                }
                catch (IOException ex)
                {
                    _log.Error($"Произошла ошибка ввода/вывода: {ex.Message}");
                }
                catch (Exception ex)
                {
                    _log.Error($"Непредвиденная ошибка при записи в CSV-файл: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                _log.Error($"Произошла ошибка: {ex.Message}");
            }
        }
    }
}
