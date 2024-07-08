using CsvHelper.Configuration;
using System;
using log4net;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using OpenQA.Selenium.Support.UI;
using CsvHelper;

namespace HTMLParser2
{
    internal class Program
    {
        private static readonly ILog _log = LogManager.GetLogger(typeof(Program));
        static void Main(string[] args)
        {
            log4net.Config.XmlConfigurator.Configure();
            string url = "https://clientportal.jse.co.za/reports/delta-option-and-structured-option-trades";
            var options = new ChromeOptions();
            options.AddArgument("headless");
            var driver = new ChromeDriver(options);

            //Создание списка объектов класса TableNode для хранения записей таблицы
            List<TableNode> tableNodes = new List<TableNode>();

            try
            {
                _log.Info("Начало загрузки страницы");
                // Открытие веб-страницы
                driver.Navigate().GoToUrl(url);
                _log.Info("Успешная загрузка страницы");

                // Ожидание загрузки таблицы
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                wait.Until(d => d.FindElement(By.Id("tblOptionsBody")));

                _log.Info($"Парсинг данных страницы");
                // Получение строк таблицы
                var rows = driver.FindElements(By.XPath("//*[@id=\"tblOptionsBody\"]/tr[position()>0]"));
                if (rows.Count == 0)
                {
                    throw new Exception("Не обнаружены строки в таблице");
                }
                _log.Info("Получены строки таблицы");

                // Перебор всех ячеек в каждой строке таблицы
                foreach (var row in rows)
                {
                    var cells = row.FindElements(By.TagName("td"));
                    try
                    {
                        tableNodes.Add(new TableNode()
                        {
                            TradeName = cells[0].Text,
                            TradeType = cells[1].Text,
                            ShortName = cells[2].Text,
                            FutureExpiry = cells[3].Text,
                            Strike = cells[4].Text,
                            CallPut = cells[5].Text,
                            Quantity = cells[6].Text,
                            Vol = cells[7].Text,
                            Premium = cells[8].Text,
                            FuturesPrice = cells[9].Text
                        });
                    }
                    catch (Exception ex)
                    {
                        _log.Error($"Ошибка извлечения данных из таблицы: {ex.Message}");
                        break;
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
            catch (WebDriverTimeoutException)
            {
                _log.Error("Не найден элемент таблицы на странице");
            }
            catch (Exception ex)
            {
                _log.Error($"Произошла ошибка: {ex.Message}");
            }
            finally
            {
                driver.Quit();
                _log.Info("Драйвер закрыт");
            }
        }
    }
}
