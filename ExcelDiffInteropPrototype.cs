using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace eWorkhelper
{
    /// <summary>
    /// M1 的最小 Excel Interop 读写原型。调用方必须在 Excel 宿主线程调用，不能放入 Task.Run。
    /// 该类不负责打开或关闭 Workbook，也不释放宿主拥有的 Application。
    /// </summary>
    internal sealed class ExcelDiffInteropPrototypeReader
    {
        internal ExcelSheetData ReadWorksheet(Excel.Worksheet worksheet)
        {
            return ReadWorksheet(worksheet, CancellationToken.None);
        }

        internal ExcelSheetData ReadWorksheet(Excel.Worksheet worksheet, CancellationToken cancellationToken)
        {
            if (worksheet == null)
            {
                throw new ArgumentNullException("worksheet");
            }

            Excel.Range cells = null;
            Excel.Range anchor = null;
            Excel.Range lastRowCell = null;
            Excel.Range lastColumnCell = null;
            Excel.Range firstDataCell = null;
            Excel.Range lastDataCell = null;
            Excel.Range dataRange = null;
            try
            {
                cancellationToken.ThrowIfCancellationRequested();
                cells = worksheet.Cells;
                anchor = cells[1, 1];
                lastRowCell = cells.Find(
                    What: "*",
                    After: anchor,
                    LookIn: Excel.XlFindLookIn.xlFormulas,
                    LookAt: Excel.XlLookAt.xlPart,
                    SearchOrder: Excel.XlSearchOrder.xlByRows,
                    SearchDirection: Excel.XlSearchDirection.xlPrevious,
                    MatchCase: false);
                lastColumnCell = cells.Find(
                    What: "*",
                    After: anchor,
                    LookIn: Excel.XlFindLookIn.xlFormulas,
                    LookAt: Excel.XlLookAt.xlPart,
                    SearchOrder: Excel.XlSearchOrder.xlByColumns,
                    SearchDirection: Excel.XlSearchDirection.xlPrevious,
                    MatchCase: false);

                if (lastRowCell == null || lastColumnCell == null)
                {
                    return new ExcelSheetData(worksheet.Name, new List<IList<ExcelCellValue>>());
                }

                int rowCount = lastRowCell.Row;
                int columnCount = lastColumnCell.Column;
                firstDataCell = cells[1, 1];
                lastDataCell = cells[rowCount, columnCount];
                dataRange = worksheet.Range[firstDataCell, lastDataCell];

                object values = dataRange.Value2;
                cancellationToken.ThrowIfCancellationRequested();
                object formulas = dataRange.Formula;
                cancellationToken.ThrowIfCancellationRequested();
                object numberFormats = dataRange.NumberFormat;
                IList<IList<ExcelCellValue>> rows = new List<IList<ExcelCellValue>>();
                for (int row = 1; row <= rowCount; row++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    IList<ExcelCellValue> convertedRow = new List<ExcelCellValue>();
                    for (int column = 1; column <= columnCount; column++)
                    {
                        if (((row - 1) * columnCount + column & 255) == 0)
                            cancellationToken.ThrowIfCancellationRequested();
                        object rawValue = GetArrayValue(values, row, column, rowCount, columnCount);
                        object formulaValue = GetArrayValue(formulas, row, column, rowCount, columnCount);
                        object numberFormatValue = GetArrayValue(numberFormats, row, column, rowCount, columnCount);
                        if (numberFormatValue == null && IsNumericValue(rawValue))
                        {
                            numberFormatValue = ReadCellNumberFormat(dataRange, row, column);
                            cancellationToken.ThrowIfCancellationRequested();
                        }
                        convertedRow.Add(ConvertCell(rawValue, formulaValue, numberFormatValue));
                    }

                    rows.Add(convertedRow);
                }

                return new ExcelSheetData(worksheet.Name, rows);
            }
            finally
            {
                ComHelper.Release(dataRange);
                ComHelper.Release(lastDataCell);
                ComHelper.Release(firstDataCell);
                ComHelper.Release(lastColumnCell);
                ComHelper.Release(lastRowCell);
                ComHelper.Release(anchor);
                ComHelper.Release(cells);
            }
        }

        private static object GetArrayValue(object values, int row, int column, int rowCount, int columnCount)
        {
            Array array = values as Array;
            if (array == null)
            {
                return rowCount == 1 && columnCount == 1 ? values : null;
            }

            return array.GetValue(row, column);
        }

        private static ExcelCellValue ConvertCell(object rawValue, object formulaValue, object numberFormatValue)
        {
            string formula = formulaValue as string;
            bool hasFormula = !string.IsNullOrEmpty(formula)
                && formula.StartsWith("=", StringComparison.Ordinal);
            string numberFormat = numberFormatValue as string ?? string.Empty;
            bool isError;
            ExcelCellSourceKind kind = !hasFormula
                ? GetSourceKind(rawValue, numberFormat, out isError)
                : ExcelCellSourceKind.Formula;
            string text = ConvertComparisonText(rawValue, numberFormat, out isError);
            return new ExcelCellValue(
                rawValue != null,
                text,
                kind,
                rawValue,
                hasFormula ? formula : string.Empty,
                text,
                numberFormat,
                isError);
        }

        private static object ReadCellNumberFormat(Excel.Range dataRange, int row, int column)
        {
            Excel.Range cells = null;
            Excel.Range cell = null;
            try
            {
                cells = dataRange.Cells;
                cell = cells[row, column] as Excel.Range;
                return cell == null ? null : cell.NumberFormat;
            }
            finally
            {
                ComHelper.Release(cell);
                ComHelper.Release(cells);
            }
        }

        private static bool IsNumericValue(object value)
        {
            return value is double || value is float || value is decimal || value is int || value is long;
        }

        private static ExcelCellSourceKind GetSourceKind(object value, string numberFormat, out bool isError)
        {
            isError = IsExcelError(value);
            if (value == null)
            {
                return ExcelCellSourceKind.Missing;
            }

            if (isError)
            {
                return ExcelCellSourceKind.Error;
            }

            if (value is string)
            {
                return string.IsNullOrEmpty((string)value)
                    ? ExcelCellSourceKind.Blank
                    : ExcelCellSourceKind.Text;
            }

            if (value is bool)
            {
                return ExcelCellSourceKind.Boolean;
            }

            if (value is DateTime)
            {
                return ExcelCellSourceKind.Date;
            }

            if (value is double || value is float || value is decimal || value is int || value is long)
            {
                return IsDateNumberFormat(numberFormat)
                    ? ExcelCellSourceKind.Date
                    : ExcelCellSourceKind.Number;
            }

            return ExcelCellSourceKind.Other;
        }

        private static string ConvertComparisonText(object value, string numberFormat, out bool isError)
        {
            isError = IsExcelError(value);
            if (value == null)
            {
                return string.Empty;
            }

            if (isError)
            {
                return ErrorText(value);
            }

            string text = value as string;
            if (text != null)
            {
                return text;
            }

            DateTime date = value is DateTime ? (DateTime)value : default(DateTime);
            if (value is DateTime)
            {
                return date.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
            }

            if (value is bool)
            {
                return (bool)value ? "TRUE" : "FALSE";
            }

            if (value is double || value is float || value is decimal || value is int || value is long)
            {
                double numericValue;
                if (double.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture), NumberStyles.Float, CultureInfo.InvariantCulture, out numericValue)
                    && IsDateNumberFormat(numberFormat)
                    && numericValue >= 0d
                    && numericValue <= 2958465d)
                {
                    try
                    {
                        return DateTime.FromOADate(numericValue).ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                    }
                    catch (ArgumentException)
                    {
                        // 按普通数字继续处理，避免仅凭格式把无效数值当作日期。
                    }
                }
            }

            IFormattable formattable = value as IFormattable;
            return formattable == null
                ? Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty
                : formattable.ToString(null, CultureInfo.InvariantCulture);
        }

        private static bool IsDateNumberFormat(string numberFormat)
        {
            if (string.IsNullOrEmpty(numberFormat) || numberFormat == "General")
            {
                return false;
            }

            string format = numberFormat.ToLowerInvariant();
            format = System.Text.RegularExpressions.Regex.Replace(format, "\"[^\"]*\"", string.Empty);
            format = System.Text.RegularExpressions.Regex.Replace(format, "\\[[^\\]]*\\]", string.Empty);
            return format.IndexOf("yy", StringComparison.Ordinal) >= 0
                || format.IndexOf("dd", StringComparison.Ordinal) >= 0
                || format.IndexOf("mm", StringComparison.Ordinal) >= 0
                || format.IndexOf("hh", StringComparison.Ordinal) >= 0
                || format.IndexOf("ss", StringComparison.Ordinal) >= 0;
        }

        private static bool IsExcelError(object value)
        {
            int code = GetExcelErrorCode(value);
            return code == 2000 || code == 2007 || code == 2015 || code == 2023
                || code == 2029 || code == 2036 || code == 2042;
        }

        private static string ErrorText(object value)
        {
            int code = GetExcelErrorCode(value);
            switch (code)
            {
                case 2000: return "#NULL!";
                case 2007: return "#DIV/0!";
                case 2015: return "#VALUE!";
                case 2023: return "#REF!";
                case 2029: return "#NAME?";
                case 2036: return "#NUM!";
                case 2042: return "#N/A";
                default: return "#ERROR!";
            }
        }

        private static int GetExcelErrorCode(object value)
        {
            int code;
            if (value is int)
            {
                code = (int)value;
            }
            else if (value is short)
            {
                code = (short)value;
            }
            else
            {
                return -1;
            }

            // Excel Interop may expose a formula error as an HRESULT such as
            // 0x800A07D7 instead of the documented xlErrDiv0 value 2007.
            if (code < 0)
            {
                code = unchecked((int)((uint)code & 0xFFFFu));
            }

            return code;
        }
    }

    /// <summary>
    /// 将结构化差异写回现有结果 Worksheet。M1 只验证坐标和值写回，不实现高亮、备注和完整格式。
    /// </summary>
    internal sealed class ExcelDiffInteropPrototypeWriter
    {
        internal void WriteDifferences(Excel.Worksheet worksheet, ExcelDiffResult result)
        {
            if (worksheet == null)
            {
                throw new ArgumentNullException("worksheet");
            }

            if (result == null)
            {
                throw new ArgumentNullException("result");
            }

            foreach (ExcelDiffCell difference in result.Differences)
            {
                Excel.Range cell = null;
                try
                {
                    cell = worksheet.Cells[difference.Coordinate.Row, difference.Coordinate.Column];
                    cell.Value2 = difference.CandidateValue.ComparisonText;
                }
                finally
                {
                    ComHelper.Release(cell);
                }
            }
        }
    }
}
