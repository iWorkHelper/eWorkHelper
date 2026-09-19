using System;
using System.Collections.Generic;

namespace eWorkhelper
{
    internal enum ExcelCellSourceKind
    {
        Missing,
        Blank,
        Text,
        Number,
        Date,
        Boolean,
        Formula,
        Error,
        Other
    }

    /// <summary>
    /// 托管比较值。ComparisonText 是参考项目 GetRows 返回的字符串语义；Exists 仅保留来源信息，
    /// 不改变“缺失单元格和空字符串均按空字符串比较”的兼容行为。
    /// </summary>
    internal sealed class ExcelCellValue
    {
        internal ExcelCellValue(bool exists, string comparisonText, ExcelCellSourceKind sourceKind, object rawValue)
            : this(exists, comparisonText, sourceKind, rawValue, null, null, null, false)
        {
        }

        internal ExcelCellValue(
            bool exists,
            string comparisonText,
            ExcelCellSourceKind sourceKind,
            object rawValue,
            string formulaText,
            string displayText,
            string numberFormat,
            bool isError)
        {
            Exists = exists;
            ComparisonText = comparisonText ?? string.Empty;
            SourceKind = sourceKind;
            RawValue = rawValue;
            FormulaText = formulaText ?? string.Empty;
            DisplayText = displayText ?? ComparisonText;
            NumberFormat = numberFormat ?? string.Empty;
            IsError = isError;
        }

        internal bool Exists { get; private set; }
        internal string ComparisonText { get; private set; }
        internal ExcelCellSourceKind SourceKind { get; private set; }
        internal object RawValue { get; private set; }
        internal string FormulaText { get; private set; }
        internal string DisplayText { get; private set; }
        internal string NumberFormat { get; private set; }
        internal bool IsError { get; private set; }

        internal static ExcelCellValue Missing()
        {
            return new ExcelCellValue(false, string.Empty, ExcelCellSourceKind.Missing, null);
        }

        internal static ExcelCellValue FromText(string text)
        {
            return new ExcelCellValue(true, text ?? string.Empty,
                string.IsNullOrEmpty(text) ? ExcelCellSourceKind.Blank : ExcelCellSourceKind.Text, text);
        }
    }

    internal sealed class ExcelSheetData
    {
        internal ExcelSheetData(string sheetName, IList<IList<ExcelCellValue>> rows)
        {
            SheetName = sheetName ?? string.Empty;
            Rows = rows ?? new List<IList<ExcelCellValue>>();
        }

        internal string SheetName { get; private set; }
        internal IList<IList<ExcelCellValue>> Rows { get; private set; }

        internal ExcelCellValue GetCell(int zeroBasedRow, int zeroBasedColumn)
        {
            if (zeroBasedRow < 0 || zeroBasedColumn < 0 || zeroBasedRow >= Rows.Count)
            {
                return ExcelCellValue.Missing();
            }

            IList<ExcelCellValue> row = Rows[zeroBasedRow];
            if (row == null || zeroBasedColumn >= row.Count || row[zeroBasedColumn] == null)
            {
                return ExcelCellValue.Missing();
            }

            return row[zeroBasedColumn];
        }

        internal static ExcelSheetData FromComparisonRows(string sheetName, params string[][] rows)
        {
            IList<IList<ExcelCellValue>> converted = new List<IList<ExcelCellValue>>();
            if (rows != null)
            {
                foreach (string[] row in rows)
                {
                    IList<ExcelCellValue> convertedRow = new List<ExcelCellValue>();
                    if (row != null)
                    {
                        foreach (string value in row)
                        {
                            convertedRow.Add(ExcelCellValue.FromText(value));
                        }
                    }

                    converted.Add(convertedRow);
                }
            }

            return new ExcelSheetData(sheetName, converted);
        }
    }

    internal struct ExcelCellCoordinate
    {
        internal ExcelCellCoordinate(int row, int column)
        {
            Row = row;
            Column = column;
        }

        internal int Row { get; private set; }
        internal int Column { get; private set; }

        public override string ToString()
        {
            return string.Format("R{0}C{1}", Row, Column);
        }
    }

    internal enum ExcelDiffKind
    {
        Added,
        Removed,
        Changed
    }

    internal enum ExcelDiffRowStatus
    {
        Unchanged,
        Added,
        Deleted,
        Changed
    }

    internal sealed class ExcelDiffRow
    {
        internal ExcelDiffRow(int rowNumber, ExcelDiffRowStatus status, int differenceCount)
        {
            RowNumber = rowNumber;
            Status = status;
            DifferenceCount = differenceCount;
        }

        internal int RowNumber { get; private set; }
        internal ExcelDiffRowStatus Status { get; private set; }
        internal int DifferenceCount { get; private set; }

        internal string StatusText
        {
            get
            {
                switch (Status)
                {
                    case ExcelDiffRowStatus.Added: return "ADD";
                    case ExcelDiffRowStatus.Deleted: return "DEL";
                    case ExcelDiffRowStatus.Changed: return "CHG";
                    default: return string.Empty;
                }
            }
        }
    }

    internal sealed class ExcelDiffCell
    {
        internal ExcelDiffCell(
            ExcelCellCoordinate coordinate,
            ExcelCellValue baselineValue,
            ExcelCellValue candidateValue,
            ExcelDiffKind kind)
        {
            Coordinate = coordinate;
            BaselineValue = baselineValue ?? ExcelCellValue.Missing();
            CandidateValue = candidateValue ?? ExcelCellValue.Missing();
            Kind = kind;
        }

        internal ExcelCellCoordinate Coordinate { get; private set; }
        internal ExcelCellValue BaselineValue { get; private set; }
        internal ExcelCellValue CandidateValue { get; private set; }
        internal ExcelDiffKind Kind { get; private set; }
    }

    internal sealed class ExcelDiffResult
    {
        internal ExcelDiffResult(
            string baselineSheetName,
            string candidateSheetName,
            int comparedRowCount,
            int comparedColumnCount,
            IList<ExcelDiffCell> differences,
            IList<ExcelDiffRow> rows)
        {
            BaselineSheetName = baselineSheetName ?? string.Empty;
            CandidateSheetName = candidateSheetName ?? string.Empty;
            ComparedRowCount = comparedRowCount;
            ComparedColumnCount = comparedColumnCount;
            Differences = differences ?? new List<ExcelDiffCell>();
            Rows = rows ?? new List<ExcelDiffRow>();
        }

        internal string BaselineSheetName { get; private set; }
        internal string CandidateSheetName { get; private set; }
        internal int ComparedRowCount { get; private set; }
        internal int ComparedColumnCount { get; private set; }
        internal IList<ExcelDiffCell> Differences { get; private set; }
        /// <summary>仅保存非 Unchanged 行，正常行在结果副本中保持 A 列空白。</summary>
        internal IList<ExcelDiffRow> Rows { get; private set; }
        internal bool HasDifferences { get { return Differences.Count != 0; } }
        internal int AddedRowCount { get { return CountRows(ExcelDiffRowStatus.Added); } }
        internal int DeletedRowCount { get { return CountRows(ExcelDiffRowStatus.Deleted); } }
        internal int ChangedRowCount { get { return CountRows(ExcelDiffRowStatus.Changed); } }

        private int CountRows(ExcelDiffRowStatus status)
        {
            int count = 0;
            foreach (ExcelDiffRow row in Rows)
            {
                if (row.Status == status) count++;
            }
            return count;
        }
    }
}
