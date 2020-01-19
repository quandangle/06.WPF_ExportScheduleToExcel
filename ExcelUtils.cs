#region Namespaces

using System;
using System.Data;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using DataTable = System.Data.DataTable;

#endregion

namespace QApps
{
    internal static class ExcelUtils
    {
        internal static DataTable ReadFile(string path, int numberColumn)
        {
            char[] tabs = { '\t' };         // tab ngang
            char[] quotes = { '\"' };       // dấu nháy kép "

            DataTable table = new DataTable("dataFromFile");


            for (int i = 0; i < numberColumn; i++)
            {
                table.Columns.Add(new DataColumn("col" + i, typeof(string)));
            }

            using (StreamReader sr = new StreamReader(path))
            {
                table.BeginLoadData();
                string line;
                //int rowsCount = 0;

                string firstLine = sr.ReadLine();
                // string otherLine = sr.ReadToEnd();

                //  string[] firstLineData = firstLine.Split(new[] { "\"", "\t" }, StringSplitOptions.RemoveEmptyEntries);

                string[] firstLineData = (
                    from s in firstLine.Split(tabs)
                    select s.Trim(quotes)).ToArray();

                if (firstLineData.Length == numberColumn)
                {
                    table.LoadDataRow(firstLineData, true);
                    //rowsCount++;
                }
                else
                {
                    foreach (string item in firstLineData)
                    {
                        if (item != String.Empty)
                        {
                            table.Rows.Add();
                            table.Rows[0][0] = item;
                            //  rowsCount++;
                            break;
                        }
                    }
                }

                while (true)
                {
                    line = sr.ReadLine();
                    if (line == null) break;

                    string[] array = (
                        from s in line.Split(tabs)
                        select s.Trim(quotes)).ToArray();

                    if (array.Length == numberColumn)
                    {
                        table.LoadDataRow(array, true);
                    }
                }

            }

            table.EndLoadData();
            return table;
        }

        internal static void ExportDataTableToExcel(ExcelPackage excelPackage, DataTable dt,
            string workSheetName, string filePath)
        {
            if (workSheetName.Length > 31)
            {
                workSheetName = workSheetName.Substring(0, 30);
            }

            // Tạo author cho file Excel
            // excelPackage.Workbook.Properties.Author = "Q'Apps SOLUTIONS";
            // excelPackage.Workbook.Properties.Comments = "This file created by Revit API from ADD-INS Q'APPS";

            // Add a new worksheet to the empty workbook
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add(workSheetName);

            // Đổ data vào Excel file
            // worksheet.Cells[1, 1].LoadFromCollection(list, true);

            #region Đổ data từ DataGridView vào Excel file

            int iRow = 1;
            int iCol;
            foreach (DataRow r in dt.Rows)
            {
                iCol = 1;
                foreach (DataColumn c in dt.Columns)
                {
                    worksheet.Cells[iRow, iCol].Value = r[c.ColumnName];
                    iCol++;
                }
                iRow++;
            }

            #endregion

            worksheet.Cells.AutoFitColumns();

            FileInfo existingFile = new FileInfo(filePath);
            excelPackage.SaveAs(existingFile);
        }
    }
}