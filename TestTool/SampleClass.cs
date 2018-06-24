using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using OfficeOpenXml;
using System.Xml;
using System.Drawing;
using OfficeOpenXml.Style;

namespace TestTool
{
    public class SampleClass
    {
        public SampleClass()
        {
        }

        private static List<String> filterColumnContent(String[] contents)
        {
            Console.WriteLine("Enter filterColumn()\n");
            if (contents == null) {
                Console.WriteLine("Input Error!\n");
                return null;
            }

            List<String> filteredColumn = new List<string>();
            foreach (String content in contents) {
                if (!filteredColumn.Contains(content)) {
                    filteredColumn.Add(content);
                    Console.WriteLine("filtered value: {0}\n", content);
                }
            }

            Console.WriteLine("Exit filterColumn()\n");
            return filteredColumn;
        }

        private static String[] getColumnContent(ExcelWorksheet sheet, String columnName)
        {
            Console.WriteLine("Enter getColumnContent(): sheet={0}, column={1}\n", sheet, columnName);

            if (sheet == null || columnName == null) {
                Console.WriteLine("Input Error!");
                return null;
            }

            int columnIndex = getColumnIndex(sheet, columnName);
            String[] content = null;

            if (columnIndex > 0) {
                int rowLength = getSheetRowLength(sheet);
                if (rowLength > 1) {
                    content = new String[rowLength - 1];
                    for (int index = 0; index < rowLength - 1; index++) {
                        content[index] = sheet.Cells[2 + index, columnIndex].Value.ToString();
                        Console.WriteLine("get value: content[{0}]={1}\n", index, content[index]);
                    }
                }
            }

            Console.WriteLine("Exit getColumnContent()\n");
            return content;
        }

        private static int getColumnIndex(ExcelWorksheet sheet, String columnName)
        {
            Console.WriteLine("Enter getColumnIndex(): sheet={0}, column={1}\n", sheet, columnName);

            if (sheet == null || columnName == null){
                Console.WriteLine("Input Error!");
                return 0;
            }

            int columnLength = getSheetColumnLength(sheet);
            int columnIndex = 0;

            for (int index = 1; index <= columnLength; index++) {
                if (sheet.Cells[1, index].Value.ToString().Equals(columnName)) {
                    Console.WriteLine("{0} is found in column[{1}]!\n", columnName, index);
                    columnIndex = index;
                    break;
                }
            }

            if (columnIndex == 0) {
                Console.WriteLine("NO COLUMN FOUND !!!\n");
            }

            Console.WriteLine("Exit getColumnIndex()\n");
            return columnIndex;
        }

        private static ExcelWorksheet getSheet(ExcelPackage package, String sheetName)
        {
            Console.WriteLine("Enter getSheet(): package={0}, sheet={1}\n", package, sheetName);

            if (package == null || sheetName == null) {
                Console.WriteLine("Input Error!\n");
                return null;
            }

            ExcelWorksheets sheets = package.Workbook.Worksheets;
            ExcelWorksheet wantedSheet = null;
            foreach (ExcelWorksheet sheet in sheets) {
                if (sheet.Name.Equals(sheetName)) {
                    Console.WriteLine("{0} is found!\n", sheetName);
                    wantedSheet = sheet;
                }
            }

            if (wantedSheet == null) {
                Console.WriteLine("NO SHEET FOUND !!!\n");
            }

            Console.WriteLine("Exit getSheet()\n");
            return wantedSheet;
        }

        public static int getSheetColumnLength(ExcelWorksheet sheet)
        {
            Console.WriteLine("Enter getSheetColumnLength(): sheet={0}\n", sheet.Name);

            int index = 1;
            while (index <= sheet.Dimension.End.Column) {
                Console.WriteLine("Cells[1, {0}]={1}\n", index, sheet.Cells[1, index].Value);
                if (sheet.Cells[1, index].Value == null) {
                    break;
                }
                index++;
            }

            Console.WriteLine("Exit getSheetColumnLength(): length={0}\n", index - 1);
            return index - 1;
        }

        public static int getSheetRowLength(ExcelWorksheet sheet)
        {
            Console.WriteLine("Enter getSheetRowLength(): sheet={0}\n", sheet.Name);

            int index = 1;
            while (index <= sheet.Dimension.End.Row) {
                Console.WriteLine("Cells[{0}, 1]={1}\n", index, sheet.Cells[index, 1].Value);
                if (sheet.Cells[index, 1].Value == null) {
                    break;
                }
                index++;
            }

            Console.WriteLine("Exit getSheetRowLength(): length={0}\n", index - 1);
            return index - 1;
        }

        public static String Run(String fileName)
        {
            Console.WriteLine("Starting running...\nReading file: {0}", fileName);
            Console.WriteLine();

            FileInfo existingFile = new FileInfo(fileName);
            using (ExcelPackage package = new ExcelPackage(existingFile)) {
                ExcelWorksheet testExecutionSheet = getSheet(package, "TestExecution");
                String[] contents = getColumnContent(testExecutionSheet, "Revision");
                List<String> columnRevision = filterColumnContent(contents);

                ExcelWorksheet testPlanSheet = getSheet(package, "TestPlan");
                int columnIndex = getColumnIndex(testPlanSheet, "Scenario");
                if (testPlanSheet != null && columnIndex > 0) {
                    Console.WriteLine("Inserting columns...\n");
                    testPlanSheet.InsertColumn(columnIndex + 1, columnRevision.Count);
                    Console.WriteLine("Columns inserted!\n");
                }
                package.SaveAs(existingFile);
            }

            Console.WriteLine("The End!\n");
            return Utils.GetFileInfo(fileName).Name;
        }
    }
}
