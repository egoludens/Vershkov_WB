using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;

namespace VershkovWB
{
    class WBReportExcel
    {
        List<IProgressReportObserver> progressReportObservers;
        Excel.Application xlApp;
        Excel.Worksheet xlWorksheet;
        int rowsInReport;
        int headerRow;
        int subHeaderRow;

        public WBReportExcel()
        {
            progressReportObservers = new List<IProgressReportObserver>();
        }

        #region Progress_Report_Notification

        public void AddProgressReportObserver(IProgressReportObserver observer)
        {
            progressReportObservers.Add(observer);
        }

        void SendProgressReportUpdate(string updateText)
        {
            foreach (IProgressReportObserver observer in progressReportObservers)
            {
                observer.NotifyAboutProgressReport(updateText);
            }
        }

        #endregion

        public bool Open(string sourceFileName)
        {
            if (String.IsNullOrWhiteSpace(sourceFileName))
            {
                SendProgressReportUpdate("Не выбран файл с отчетом Wildberries.");
                return false;
            }

            if (!System.IO.File.Exists(sourceFileName))
            {
                SendProgressReportUpdate("Файл с отчетом Wildberries не найден. Проверьте, что он существует, пожалуйста.");
                return false;
            }

            SendProgressReportUpdate("Запускаем Excel...");

            Excel.Workbook xlSourceWorkbook;
            try
            {
                xlApp = new Excel.Application();
                xlSourceWorkbook = xlApp.Workbooks.Open(sourceFileName);
            }
            catch
            {
                SendProgressReportUpdate("Не удалось открыть файл в Excel. Проверьте следующее:\r\n1. У вас должен быть установлен Excel.\r\n2. Файл с отчетом WB должен открываться в Excel.");
                return false;
            }

            Excel.Worksheet xlSourceWorksheet = xlSourceWorkbook.Sheets[1];
            xlSourceWorksheet.Copy();
            xlWorksheet = xlApp.ActiveWorkbook.Sheets[1];
            xlWorksheet.Name = "Отчет";
            xlSourceWorkbook.Close();

            rowsInReport = xlWorksheet.UsedRange.Rows.Count;
            headerRow = 1;
            subHeaderRow = 2;

            return true;
        }

        public void SetVisibility(bool isVisible)
        {
            xlApp.Visible = isVisible;
        }

        public bool RemoveColumnWithSubHeader(string subHeaderText)
        {
            SendProgressReportUpdate("Удаляем колонку с подзаголовком \"" + subHeaderText + "\"...");
            int targetColumn = FindColumnNumberBySubHeader(subHeaderText);

            if (targetColumn == 0)
            {
                SendProgressReportUpdate("Не найдена колонка с подзаголовком \"" + subHeaderText + "\", чтобы её удалить.");
                return false;
            }

            return removeColumnByNumber(targetColumn);
        }

        public bool RemoveColumnWithHeaderAndSubHeader(string headerText, string subHeaderText)
        {
            SendProgressReportUpdate("Удаляем колонку с заголовком \"" + headerText + "\" и подзаголовком \"" + subHeaderText + "\"...");
            int targetColumn = FindColumnNumberByHeaderAndSubHeader(headerText, subHeaderText);

            if (targetColumn == 0)
            {
                SendProgressReportUpdate("Не найдена колонка с заголовком \"" + headerText + "\" и подзаголовком \"" + subHeaderText + "\", чтобы её удалить.");
                return false;
            }

            return removeColumnByNumber(targetColumn);
        }

        string getTextFromCellConsideringPossibleMerge(int row, int column)
        {
            Excel.Range xlCell = xlWorksheet.Cells[row, column];
            try
            {
                if (xlCell.MergeCells)
                {
                    return xlCell.MergeArea.Cells[1, 1].Value.ToString();
                }
                else
                {
                    return xlCell.Value.ToString();
                }
            }
            catch
            {
                SendProgressReportUpdate("Ошибка при чтении текста из ячейки в строке " + row.ToString() + ", колонке " + column.ToString() + "!");
                return "";
            }

        }

        int getIntFromCellConsideringPossibleMerge(Excel.Worksheet xlWS, int row, int column)
        {
            try
            {
                Excel.Range xlCell = xlWS.Cells[row, column];
                if (xlCell.MergeCells)
                {
                    return (int)xlCell.MergeArea.Cells[1, 1].Value;
                }
                else
                {
                    return (int)xlCell.Value;
                }
            }
            catch
            {
                SendProgressReportUpdate("Ошибка при чтении числа из ячейки в строке " + row.ToString() + ", колонке " + column.ToString() + "!");
                return 0;
            }

        }

        bool removeColumnByNumber(int targetColumn)
        {
            Excel.Range xlColumns = xlWorksheet.Columns;
            Excel.Range xlColumn = (Excel.Range)xlColumns[targetColumn, System.Reflection.Missing.Value];
            xlColumn.Delete();

            return true;
        }

        bool removeRowByNumber(int targetRow)
        {
            Excel.Range xlRows = xlWorksheet.Rows;
            Excel.Range xlRow = (Excel.Range)xlRows[targetRow, System.Reflection.Missing.Value];
            xlRow.Delete();

            return true;
        }

        public int FindColumnNumberBySubHeader(string subHeaderText)
        {
            int columnsInReport = xlWorksheet.UsedRange.Columns.Count;
            int targetColumn = 0;
            for (int i = 1; i <= columnsInReport; i++)
            {
                string subHeaderCellText = getTextFromCellConsideringPossibleMerge(subHeaderRow, i).Trim();
                if (subHeaderCellText == subHeaderText.Trim())
                {
                    targetColumn = i;
                    break;
                }
            }
            return targetColumn;
        }

        public int FindColumnNumberByHeaderAndSubHeader(string headerText, string subHeaderText)
        {
            int columnsInReport = xlWorksheet.UsedRange.Columns.Count;
            int targetColumn = 0;
            for (int i = 1; i <= columnsInReport; i++)
            {
                string headerCellText = getTextFromCellConsideringPossibleMerge(headerRow, i).Trim();
                string subHeaderCellText = getTextFromCellConsideringPossibleMerge(subHeaderRow, i).Trim();

                if (headerCellText == headerText.Trim() && subHeaderCellText == subHeaderText.Trim())
                {
                    targetColumn = i;
                    break;
                }
            }
            return targetColumn;
        }

        public void RollUp(int keyColumnNumber, int[] valueColumnsNumbers)
        {
            SendProgressReportUpdate("Начата свертка отчета по колонке №" + keyColumnNumber.ToString() + "...");
            Dictionary<string, int> rowNumbersForKeys = new Dictionary<string, int>();
            int rowsProcessed = 0;
            for (int i = subHeaderRow + 1; i <= rowsInReport; i++) // Идем со строки после заголовка
            {
                string keyValue = getTextFromCellConsideringPossibleMerge(i, keyColumnNumber);
                int firstRowForKey = 0;
                if (rowNumbersForKeys.TryGetValue(keyValue, out firstRowForKey))
                {
                    // Добавляем значения к ранее обработанной строке, удаляем текущую строку, вносим изменения в итерационные переменные
                    foreach (int valueColumnNumber in valueColumnsNumbers)
                    {
                        int valueInFirstRow = getIntFromCellConsideringPossibleMerge(xlWorksheet, firstRowForKey, valueColumnNumber);
                        int valueInThisRow = getIntFromCellConsideringPossibleMerge(xlWorksheet, i, valueColumnNumber);
                        xlWorksheet.Cells[firstRowForKey, valueColumnNumber].Value = valueInFirstRow + valueInThisRow;
                    }
                    removeRowByNumber(i);
                    i--;
                    rowsInReport--;
                }
                else
                {
                    // Запоминаем строку как обработанную
                    rowNumbersForKeys.Add(keyValue, i);
                }

                rowsProcessed++;
                if (rowsProcessed % 100 == 0)
                {
                    SendProgressReportUpdate("Обработано " + rowsProcessed.ToString() + " строк...");
                }
            }
            SendProgressReportUpdate("Cвертка отчета по колонке №" + keyColumnNumber.ToString() + " закончена.");
        }

        public void AddColumnByTemplate(string headerText, string subHeaderText, string template, int templateInputColumnNumber, int colWidth)
        {
            SendProgressReportUpdate("Добавляем колонку с заголовком \"" + headerText + "\" и подзаголовком \"" + subHeaderText + "\"...");

            int newColumnNumber = FindColumnNumberByHeaderAndSubHeader("", "");
            if (newColumnNumber == 0)
            {
                newColumnNumber = xlWorksheet.UsedRange.Columns.Count + 1;
            }
            xlWorksheet.Cells[headerRow, newColumnNumber].Value = headerText;
            xlWorksheet.Cells[subHeaderRow, newColumnNumber].Value = subHeaderText;

            xlWorksheet.Cells[1, newColumnNumber].EntireColumn.ColumnWidth = colWidth;

            int rowsInReport = xlWorksheet.UsedRange.Rows.Count;
            for (int i = 3; i <= rowsInReport; i++) // Идем с третьей строки, т.к. на первых двух у нас заголовок
            {
                string inputValueForTemplate = getTextFromCellConsideringPossibleMerge(i, templateInputColumnNumber);
                if (!String.IsNullOrWhiteSpace(inputValueForTemplate))
                {
                    string newValue = template.Replace("%1", inputValueForTemplate);
                    Excel.Range xlCell = xlWorksheet.Cells[i, newColumnNumber];
                    xlCell.Value = newValue;
                    xlWorksheet.Hyperlinks.Add(xlCell, newValue);
                }
            }
        }

        public void AddSubtotals(int[] columnNumbers)
        {
            SendProgressReportUpdate("Добавляем подытоги...");
            Excel.Range xlFirstRow = xlWorksheet.get_Range("A1", "A1").EntireRow;
            xlFirstRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
            rowsInReport++;
            headerRow++;
            subHeaderRow++;
            foreach (int columnNumber in columnNumbers)
            {
                Excel.Range xlCell = xlWorksheet.Cells[1, columnNumber];
                string cellAddress = xlCell.Address.Replace("$", "");
                string subtotalTopCellAddress = cellAddress.Replace("1", "4");
                string subtotalBottomCellAddress = cellAddress.Replace("1", rowsInReport.ToString());
                string formula = "ПРОМЕЖУТОЧНЫЕ.ИТОГИ(9;%1:%2)".Replace("%1", subtotalTopCellAddress).Replace("%2", subtotalBottomCellAddress);

                xlCell.Value = formula;
                xlCell.Replace(formula, "=" + formula);
            }
        }

        public void SetAutoFilterOnSubHeader()
        {
            SendProgressReportUpdate("Добавляем фильтр...");

            if (xlWorksheet.AutoFilter != null)
            {
                xlWorksheet.AutoFilterMode = false;
            }

            Excel.Range xlFilterRow = xlWorksheet.Cells[subHeaderRow, 1].EntireRow;
            xlFilterRow.AutoFilter(1);
        }

        public string GetReportTitle()
        {
            return getTextFromCellConsideringPossibleMerge(headerRow, 1);
        }

        public void InitiateSaveAs(string offeredFileName)
        {
            SendProgressReportUpdate("Предлагаем сохранить файл как \"" + offeredFileName + "\"...");
            var fileName = xlApp.GetSaveAsFilename(offeredFileName);
            if (fileName is String)
            {
                SendProgressReportUpdate("Cохраняем файл как \"" + fileName + "\"...");
                xlApp.ActiveWorkbook.SaveAs(fileName);
            }
        }

        public void CreateSubtotalsWorksheet(string worksheetName, int keyColumn, int[] valueColumns)
        {
            SendProgressReportUpdate("Формируем лист \"" + worksheetName + "\"...");
            Excel.Worksheet xlWorksheetSubtotals = xlApp.ActiveWorkbook.Sheets.Add(System.Reflection.Missing.Value, xlWorksheet);
            xlWorksheetSubtotals.Name = worksheetName;

            // Формирование заголовка
            xlWorksheetSubtotals.Cells[1, 1].Value = getTextFromCellConsideringPossibleMerge(headerRow, keyColumn);
            xlWorksheetSubtotals.Cells[2, 1].Value = getTextFromCellConsideringPossibleMerge(subHeaderRow, keyColumn);
            int subtotalColumn = 1;
            foreach (int valueColumn in valueColumns)
            {
                subtotalColumn++;
                xlWorksheetSubtotals.Cells[1, subtotalColumn].Value = getTextFromCellConsideringPossibleMerge(headerRow, valueColumn);
                xlWorksheetSubtotals.Cells[2, subtotalColumn].Value = getTextFromCellConsideringPossibleMerge(subHeaderRow, valueColumn);
            }

            Dictionary<string, int> rowNumbersForKeys = new Dictionary<string, int>();
            int rowsProcessed = 0;
            int subtotalRows = 0;

            for (int i = subHeaderRow + 1; i <= rowsInReport; i++) // Идем со строки после заголовка
            {
                string keyValue = getTextFromCellConsideringPossibleMerge(i, keyColumn);
                int subtotalRowForKey = 0;
                if (!rowNumbersForKeys.TryGetValue(keyValue, out subtotalRowForKey))
                {
                    // Заводим новую строку итогов
                    subtotalRows++;
                    subtotalRowForKey = subtotalRows + 2;
                    xlWorksheetSubtotals.Cells[subtotalRowForKey, 1].Value = keyValue;
                    rowNumbersForKeys.Add(keyValue, subtotalRowForKey);
                    subtotalColumn = 1;
                    foreach (int valueColumn in valueColumns)
                    {
                        subtotalColumn++;
                        xlWorksheetSubtotals.Cells[subtotalRowForKey, subtotalColumn].Value = 0;
                    }
                }

                // Добавляем значения к строке итогов
                subtotalColumn = 1;
                foreach (int valueColumn in valueColumns)
                {
                    subtotalColumn++;
                    int valueInSubtotal = getIntFromCellConsideringPossibleMerge(xlWorksheetSubtotals, subtotalRowForKey, subtotalColumn);
                    int valueInReport   = getIntFromCellConsideringPossibleMerge(xlWorksheet, i, valueColumn);
                    xlWorksheetSubtotals.Cells[subtotalRowForKey, subtotalColumn].Value = valueInSubtotal + valueInReport;
                }

                rowsProcessed++;
                if (rowsProcessed % 100 == 0)
                {
                    SendProgressReportUpdate("Обработано " + rowsProcessed.ToString() + " строк...");
                }
            }

            // Формирование итогов
            int rowTotals = 3 + subtotalRows;
            xlWorksheetSubtotals.Cells[rowTotals, 1].Value = "ИТОГО:";
            subtotalColumn = 1;
            foreach (int valueColumn in valueColumns)
            {
                subtotalColumn++;

                Excel.Range xlCell = xlWorksheetSubtotals.Cells[1, subtotalColumn];
                string cellAddress = xlCell.Address.Replace("$", "");
                string dataTopCellAddress = cellAddress.Replace("1", "3");
                string dataBottomCellAddress = cellAddress.Replace("1", (2 + subtotalRows).ToString());
                string formula = "СУММ(%1:%2)".Replace("%1", dataTopCellAddress).Replace("%2", dataBottomCellAddress);

                xlWorksheetSubtotals.Cells[rowTotals, subtotalColumn].Formula = formula;
                xlWorksheetSubtotals.Cells[rowTotals, subtotalColumn].Replace(formula, "=" + formula);
            }

            // Оформляем шапку таблицы
            string tableHatTopLeft = xlWorksheetSubtotals.Cells[1, 1].Address.Replace("$", "");
            string tableHatBottomRight = xlWorksheetSubtotals.Cells[2, subtotalColumn].Address.Replace("$", "");
            Excel.Range tableHat = xlWorksheetSubtotals.get_Range(tableHatTopLeft, tableHatBottomRight);
            tableHat.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
            tableHat.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            tableHat.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            tableHat.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            tableHat.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            tableHat.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            tableHat.Columns.AutoFit();

            // Оформляем данные в таблице
            string tableDataTopLeft = xlWorksheetSubtotals.Cells[3, 1].Address.Replace("$", "");
            string tableDataBottomRight = xlWorksheetSubtotals.Cells[2 + subtotalRows, subtotalColumn].Address.Replace("$", "");
            Excel.Range tableData = xlWorksheetSubtotals.get_Range(tableDataTopLeft, tableDataBottomRight);
            tableData.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            tableData.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            tableData.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            tableData.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            tableData.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
            tableData.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

            // Оформляем подвал таблицы
            string cellarHatTopLeft = xlWorksheetSubtotals.Cells[rowTotals, 1].Address.Replace("$", "");
            string cellarHatBottomRight = xlWorksheetSubtotals.Cells[rowTotals, subtotalColumn].Address.Replace("$", "");
            Excel.Range tableCellar = xlWorksheetSubtotals.get_Range(cellarHatTopLeft, cellarHatBottomRight);
            tableCellar.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            tableCellar.Font.Bold = true;
            tableCellar.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            tableCellar.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            tableCellar.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            tableCellar.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            tableCellar.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

        }

    }
}
