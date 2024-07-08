using CsvHelper.Configuration.Attributes;
using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HTMLParser1
{
    public class TableNode
    {
        public string Name { get; set; }
        public string Last { get; set; }
        [Name("Chg. % 1D Chg. Abs.")]
        public string Changes { get; set; }
        private string dateTime;
        [Name("Date Time")]
        public string DateTime
        {
            get { return dateTime; }
            set
            {
                dateTime = FormatDateTime(value);
            }
        }
        public string ISIN { get; set; }
        private string bidVolume;
        [Name("Bid Volume")]
        public string BidVolume
        {
            get { return bidVolume; }
            set
            {
                bidVolume = GetMultiCellInnerText(value);
            }
        }
        private string askVolume;
        [Name("Ask Volume")]
        public string AskVolume
        {
            get { return askVolume; }
            set
            {
                askVolume = GetMultiCellInnerText(value);
            }
        }
        public string Maturity { get; set; }
        public string Status { get; set; }

        //Метод для форматирования строковой записи даты и времени (добавление пробела между датой и временем)
        private string FormatDateTime(string dateTime)
        {
            if (System.DateTime.TryParseExact(dateTime, "MM/dd/yyyyHH:mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.None, out System.DateTime parsedDateTime))
            {
                return parsedDateTime.ToString("MM/dd/yyyy HH:mm:ss", CultureInfo.InvariantCulture);
            }
            return dateTime;
        }
        //Метод для форматирования строковой записи данных колонок Bid Volume и Ask Volume (добавление пробела первым и вторым div в ячейке таблицы)
        private string GetMultiCellInnerText(string html)
        {
            if (html != null)
            {
                var doc = new HtmlDocument();
                doc.LoadHtml(html);
                var divs = doc.DocumentNode.SelectNodes(".//div");
                if (divs != null && divs.Count >= 2)
                {
                    return $"{divs[0].InnerText} {divs[1].InnerText}";
                }
            }
            return string.Empty;
        }

    }
}

