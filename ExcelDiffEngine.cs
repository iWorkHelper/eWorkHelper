using System;
using System.Collections.Generic;
using System.Threading;

namespace eWorkhelper
{
    /// <summary>
    /// 参考 diff-excel 的纯托管坐标比较器。
    /// 不持有 Excel COM、WinForms 或 Ribbon 对象。
    /// </summary>
    internal sealed class ExcelDiffEngine
    {
        internal ExcelDiffResult Compare(ExcelSheetData baseline, ExcelSheetData candidate)
        {
            return Compare(baseline, candidate, CancellationToken.None, null);
        }

        internal ExcelDiffResult Compare(ExcelSheetData baseline, ExcelSheetData candidate, CancellationToken cancellationToken, Action<long> progress)
        {
            if (baseline == null)
            {
                throw new ArgumentNullException("baseline");
            }

            if (candidate == null)
            {
                throw new ArgumentNullException("candidate");
            }

            int rowCount = Math.Max(baseline.Rows.Count, candidate.Rows.Count);
            int maxColumnCount = 0;
            IList<ExcelDiffCell> differences = new List<ExcelDiffCell>();
            IList<ExcelDiffRow> rows = new List<ExcelDiffRow>();
            long processed = 0;

            for (int row = 0; row < rowCount; row++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                int baselineColumnCount = GetColumnCount(baseline, row);
                int candidateColumnCount = GetColumnCount(candidate, row);
                int columnCount = Math.Max(baselineColumnCount, candidateColumnCount);
                maxColumnCount = Math.Max(maxColumnCount, columnCount);
                int rowDifferenceCount = 0;
                bool baselineHasContent = RowHasContent(baseline, row, cancellationToken);
                bool candidateHasContent = RowHasContent(candidate, row, cancellationToken);

                for (int column = 0; column < columnCount; column++)
                {
                    if ((column & 255) == 0) cancellationToken.ThrowIfCancellationRequested();
                    ExcelCellValue baselineValue = baseline.GetCell(row, column);
                    ExcelCellValue candidateValue = candidate.GetCell(row, column);
                    processed++;
                    if (progress != null && (processed & 255) == 0) progress(processed);
                    if (string.Equals(
                        baselineValue.ComparisonText,
                        candidateValue.ComparisonText,
                        StringComparison.Ordinal))
                    {
                        continue;
                    }

                    differences.Add(new ExcelDiffCell(
                        new ExcelCellCoordinate(row + 1, column + 1),
                        baselineValue,
                        candidateValue,
                        GetDiffKind(baselineValue, candidateValue)));
                    rowDifferenceCount++;
                }


                ExcelDiffRowStatus status = ExcelDiffRowStatus.Unchanged;
                if (!baselineHasContent && candidateHasContent) status = ExcelDiffRowStatus.Added;
                else if (baselineHasContent && !candidateHasContent) status = ExcelDiffRowStatus.Deleted;
                else if (baselineHasContent && candidateHasContent && rowDifferenceCount > 0) status = ExcelDiffRowStatus.Changed;
                if (status != ExcelDiffRowStatus.Unchanged)
                    rows.Add(new ExcelDiffRow(row + 1, status, rowDifferenceCount));
            }

            return new ExcelDiffResult(
                baseline.SheetName,
                candidate.SheetName,
                rowCount,
                maxColumnCount,
                differences,
                rows);
        }

        private static bool RowHasContent(ExcelSheetData sheet, int row, CancellationToken cancellationToken)
        {
            int columnCount = GetColumnCount(sheet, row);
            for (int column = 0; column < columnCount; column++)
            {
                if ((column & 255) == 0) cancellationToken.ThrowIfCancellationRequested();
                if (!string.IsNullOrEmpty(sheet.GetCell(row, column).ComparisonText)) return true;
            }
            return false;
        }

        private static int GetColumnCount(ExcelSheetData sheet, int row)
        {
            if (row < 0 || row >= sheet.Rows.Count || sheet.Rows[row] == null)
            {
                return 0;
            }

            return sheet.Rows[row].Count;
        }

        private static ExcelDiffKind GetDiffKind(ExcelCellValue baseline, ExcelCellValue candidate)
        {
            bool baselineEmpty = string.IsNullOrEmpty(baseline.ComparisonText);
            bool candidateEmpty = string.IsNullOrEmpty(candidate.ComparisonText);
            if (baselineEmpty && !candidateEmpty)
            {
                return ExcelDiffKind.Added;
            }

            if (!baselineEmpty && candidateEmpty)
            {
                return ExcelDiffKind.Removed;
            }

            return ExcelDiffKind.Changed;
        }
    }
}
