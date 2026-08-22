using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;

namespace eWorkhelper
{
    internal sealed class UnmergeAndFillService
    {
        internal int Execute(Excel.Application application)
        {
            if (application == null)
            {
                throw new InvalidOperationException("无法获取当前 Excel 应用程序。");
            }

            Excel.Range selection = application.Selection as Excel.Range;
            if (selection == null)
            {
                throw new InvalidOperationException("请先选择要处理的单元格区域。");
            }

            IList<Excel.Range> mergeAreas = FindMergeAreas(selection);
            if (mergeAreas.Count == 0)
            {
                return 0;
            }

            bool originalScreenUpdating = application.ScreenUpdating;
            bool originalEnableEvents = application.EnableEvents;

            try
            {
                application.ScreenUpdating = false;
                application.EnableEvents = false;

                foreach (Excel.Range mergeArea in mergeAreas)
                {
                    UnmergeAndFill(mergeArea);
                }

                return mergeAreas.Count;
            }
            finally
            {
                try
                {
                    application.EnableEvents = originalEnableEvents;
                }
                finally
                {
                    application.ScreenUpdating = originalScreenUpdating;
                }
            }
        }

        private static IList<Excel.Range> FindMergeAreas(Excel.Range selection)
        {
            List<Excel.Range> mergeAreas = new List<Excel.Range>();
            HashSet<string> processedKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (Excel.Range area in selection.Areas)
            {
                FindMergeAreasInRange(area, processedKeys, mergeAreas);
            }

            return mergeAreas;
        }

        private static void FindMergeAreasInRange(
            Excel.Range range,
            ISet<string> processedKeys,
            IList<Excel.Range> mergeAreas)
        {
            object mergeState = range.MergeCells;
            if (mergeState is bool && !(bool)mergeState)
            {
                return;
            }

            Excel.Range firstCell = range.Cells[1, 1] as Excel.Range;
            if (firstCell == null)
            {
                return;
            }

            if ((bool)firstCell.MergeCells)
            {
                Excel.Range firstMergeArea = firstCell.MergeArea;
                if (ContainsRange(firstMergeArea, range))
                {
                    AddMergeArea(firstMergeArea, processedKeys, mergeAreas);
                    return;
                }
            }

            int rowCount = range.Rows.Count;
            int columnCount = range.Columns.Count;
            if (rowCount == 1 && columnCount == 1)
            {
                if ((bool)range.MergeCells)
                {
                    AddMergeArea(range.MergeArea, processedKeys, mergeAreas);
                }

                return;
            }

            if (rowCount >= columnCount && rowCount > 1)
            {
                int firstRowCount = rowCount / 2;
                Excel.Range firstHalf = range.Resize[firstRowCount, columnCount];
                Excel.Range secondHalf = range.Offset[firstRowCount, 0].Resize[rowCount - firstRowCount, columnCount];
                FindMergeAreasInRange(firstHalf, processedKeys, mergeAreas);
                FindMergeAreasInRange(secondHalf, processedKeys, mergeAreas);
                return;
            }

            int firstColumnCount = columnCount / 2;
            Excel.Range leftHalf = range.Resize[rowCount, firstColumnCount];
            Excel.Range rightHalf = range.Offset[0, firstColumnCount].Resize[rowCount, columnCount - firstColumnCount];
            FindMergeAreasInRange(leftHalf, processedKeys, mergeAreas);
            FindMergeAreasInRange(rightHalf, processedKeys, mergeAreas);
        }

        private static void AddMergeArea(
            Excel.Range mergeArea,
            ISet<string> processedKeys,
            IList<Excel.Range> mergeAreas)
        {
            string key = mergeArea.get_Address(
                true,
                true,
                Excel.XlReferenceStyle.xlA1,
                true,
                Type.Missing);

            if (processedKeys.Add(key))
            {
                mergeAreas.Add(mergeArea);
            }
        }

        private static bool ContainsRange(Excel.Range container, Excel.Range candidate)
        {
            int containerLastRow = container.Row + container.Rows.Count - 1;
            int containerLastColumn = container.Column + container.Columns.Count - 1;
            int candidateLastRow = candidate.Row + candidate.Rows.Count - 1;
            int candidateLastColumn = candidate.Column + candidate.Columns.Count - 1;

            return candidate.Row >= container.Row
                && candidate.Column >= container.Column
                && candidateLastRow <= containerLastRow
                && candidateLastColumn <= containerLastColumn;
        }

        private static void UnmergeAndFill(Excel.Range mergeArea)
        {
            Excel.Range topLeft = mergeArea.Cells[1, 1] as Excel.Range;
            if (topLeft == null)
            {
                throw new InvalidOperationException("无法读取合并区域的左上角单元格。");
            }

            bool hasFormula = (bool)topLeft.HasFormula;
            object content = hasFormula ? topLeft.FormulaR1C1 : topLeft.Value2;

            mergeArea.UnMerge();

            if (content == null)
            {
                return;
            }

            if (hasFormula)
            {
                mergeArea.FormulaR1C1 = content;
            }
            else
            {
                mergeArea.Value2 = content;
            }
        }
    }
}
