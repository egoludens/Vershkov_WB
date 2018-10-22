using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;

namespace VershkovWB
{
    class WBReportProcessor : IProgressReportObserver
    {
        string sourceFileName;
        List<IProgressReportObserver> progressReportObservers;
        struct CharCodeRange { public int MinCode; public int MaxCode; }

        public WBReportProcessor(string sourceFileName)
        {
            this.sourceFileName = sourceFileName;
            progressReportObservers = new List<IProgressReportObserver>();
        }

        #region Progress_Report_Notification

        public void AddProgressReportObserver(IProgressReportObserver observer)
        {
            progressReportObservers.Add(observer);
        }

        void SendProgressReportUpdate(string updateText)
        {
            foreach(IProgressReportObserver observer in progressReportObservers)
            {
                observer.NotifyAboutProgressReport(updateText);
            }
        }

        #endregion

        #region Observer_for_Progress_Reports
        public void NotifyAboutProgressReport(string progressReportUpdate)
        {
            SendProgressReportUpdate(progressReportUpdate);
        }
        #endregion

        public void Process()
        {

            WBReportExcel reportExcel = new WBReportExcel();
            reportExcel.AddProgressReportObserver(this);
            if (reportExcel.Open(sourceFileName) == false)
            {
                return;
            }

            bool debugMode = false;

            if (debugMode)
            {
                reportExcel.SetVisibility(true);
                ProcessExcelReport(reportExcel);
                SendProgressReportUpdate("Обработка файла завершилась успешно.");
            }
            else
            {
                try
                {
                    ProcessExcelReport(reportExcel);
                    SendProgressReportUpdate("Обработка файла завершилась успешно.");
                }
                catch
                {
                    SendProgressReportUpdate("Обработка файла завершилась с ошибкой!");
                }
                reportExcel.SetVisibility(true);
            }

        }

        void ProcessExcelReport(WBReportExcel reportExcel)
        {
            reportExcel.RemoveColumnWithSubHeader("Баркод");
            reportExcel.RemoveColumnWithSubHeader("Размер");
            reportExcel.RemoveColumnWithSubHeader("Контракт");

            reportExcel.RemoveColumnWithHeaderAndSubHeader("Поступления", "с-с, руб");
            reportExcel.RemoveColumnWithHeaderAndSubHeader("Поступления", "шт");

            reportExcel.RemoveColumnWithHeaderAndSubHeader("Заказано", "с-с, руб");

            reportExcel.RemoveColumnWithHeaderAndSubHeader("Максимально", "заказано шт в день");

            reportExcel.RemoveColumnWithHeaderAndSubHeader("Возвраты до оплаты", "с-с, руб");
            reportExcel.RemoveColumnWithHeaderAndSubHeader("Возвраты до оплаты", "шт");

            reportExcel.RemoveColumnWithHeaderAndSubHeader("Продажи по оплатам", "с-с + вознагр, руб");

            reportExcel.RemoveColumnWithHeaderAndSubHeader("Возвраты со склада поставщику", "с-с, руб");
            reportExcel.RemoveColumnWithHeaderAndSubHeader("Возвраты со склада поставщику", "шт");

            int keyColumn = reportExcel.FindColumnNumberBySubHeader("Артикул поставщика");
            int[] valueColumns = {
                reportExcel.FindColumnNumberByHeaderAndSubHeader("Заказано", "шт"),
                reportExcel.FindColumnNumberByHeaderAndSubHeader("Продажи по оплатам", "шт"),
                reportExcel.FindColumnNumberByHeaderAndSubHeader("Текущий", "остаток, шт")
            };
            reportExcel.RollUp(keyColumn, valueColumns);

            int inputForTemplateColumn = reportExcel.FindColumnNumberBySubHeader("Номенклатура (код 1С)");
            reportExcel.AddColumnByTemplate("", "Ссылка", "https://www.wildberries.ru/catalog/%1/detail.aspx", inputForTemplateColumn, 60);

            reportExcel.AddSubtotals(valueColumns);
            reportExcel.SetAutoFilterOnSubHeader();


            int brandColumn = reportExcel.FindColumnNumberBySubHeader("Бренд");
            reportExcel.CreateSubtotalsWorksheet("Итоги по брендам", brandColumn, valueColumns);


            reportExcel.InitiateSaveAs(BuildFileNameFromReportTitle(reportExcel.GetReportTitle()));

        }

        string BuildFileNameFromReportTitle(string title)
        {

            CharCodeRange[] acceptableCharCodeRanges =
            {
                new CharCodeRange { MinCode = (int)'a', MaxCode = (int)'z' },
                new CharCodeRange { MinCode = (int)'A', MaxCode = (int)'Z' },
                new CharCodeRange { MinCode = (int)'а', MaxCode = (int)'я' },
                new CharCodeRange { MinCode = (int)'А', MaxCode = (int)'Я' },
                new CharCodeRange { MinCode = (int)'0', MaxCode = (int)'9' },
                new CharCodeRange { MinCode = (int)' ', MaxCode = (int)' ' },
                new CharCodeRange { MinCode = (int)'.', MaxCode = (int)'.' },
                new CharCodeRange { MinCode = (int)'ё', MaxCode = (int)'ё' }
            };

            string processedTitle = "";

            for(int i = 0; i < title.Length; i++)
            {
                int charCode = (int)title[i];
                bool charInAcceptableRange = false;
                foreach (CharCodeRange acceptableCharCodeRange in acceptableCharCodeRanges)
                {
                    if (charCode >= acceptableCharCodeRange.MinCode && charCode <= acceptableCharCodeRange.MaxCode)
                    {
                        charInAcceptableRange = true;
                        break;
                    }
                }
                if (charInAcceptableRange)
                {
                    processedTitle = processedTitle + title[i];
                }
            }

            processedTitle = processedTitle.Replace("Общество с ограниченной ответственностью", "");
            int cutoffPosition = processedTitle.IndexOf("сформирован");
            if (cutoffPosition > 0)
            {
                processedTitle = processedTitle.Substring(0, cutoffPosition);
            }
            processedTitle = processedTitle.Trim().Replace("  ", " ");
            processedTitle += ".xls";

            return processedTitle;
        }

    }
}
