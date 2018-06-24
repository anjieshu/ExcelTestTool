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

        private static List<String> filterColumn(String[] column)
        {
            Console.WriteLine("Enter filterColumn()\n");
            if (column == null || column.Length < 0) {
                Console.WriteLine("Input Error!\n");
                return null;
            }

            List<String> columnRevision = new List<string>();
            foreach (String name in column) {
                if (!columnRevision.Contains(name)) {
                    columnRevision.Add(name);
                    Console.WriteLine("Add value: {0}\n", name);
                }
            }

            Console.WriteLine("Exit filterColumn()\n");
            return columnRevision;
        }

        private static String[] findColumn(ExcelWorksheet sheet, String columnName)
        {
            Console.WriteLine("Enter findColumn(): sheet={0}, column={1}\n", sheet, columnName);

            if (sheet == null || columnName == null) {
                Console.WriteLine("Input Error!");
                return null;
            }

            int count = 0;
            for (int index = 1; index <= sheet.Dimension.End.Column; index++) {
                if (sheet.Cells[1, index].Value.ToString().Equals(columnName)) {
                    Console.WriteLine("{0} is found in column[{1}]!\n", columnName, index);
                    count = index;
                    break;
                }
            }

            String[] content = null;
            if (count != 0) {
                int rowCount = sheet.Dimension.End.Row;
                Console.WriteLine("rowCount={0}\n", rowCount);

                content = new String[8];
                for (int index = 0; index < 8; index++) {
                    content[index] = sheet.Cells[index + 2, count].Value.ToString();
                    Console.WriteLine("Add value: column[{0}]={1}\n", index, content[index]);
                }
            }

            Console.WriteLine("Exit findColumn()\n");
            return content;
        }

        private static int findColumnIndex(ExcelWorksheet sheet, String columnName)
        {
            Console.WriteLine("Enter findColumnIndex(): sheet={0}, column={1}\n", sheet, columnName);

            if (sheet == null || columnName == null){
                Console.WriteLine("Input Error!");
                return 0;
            }

            int columnCount = sheet.Dimension.End.Column;
            Console.WriteLine("columnCount={0}\n", columnCount);

            int count = 0;
            for (int index = 1; index <= sheet.Dimension.End.Column; index++) {
                if (sheet.Cells[2, index].Value.ToString().Equals(columnName)) {
                    Console.WriteLine("{0} is found in column[{1}]!\n", columnName, index);
                    count = index;
                    break;
                }
            }

            Console.WriteLine("Exit findColumnIndex()\n");
            return count;
        }

        private static ExcelWorksheet findSheet(ExcelPackage package, String sheetName)
        {
            Console.WriteLine("Enter findSheet(): package={0}, sheet={1}\n", package, sheetName);

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

            Console.WriteLine("Exit findSheet()\n");
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
            while (index <= sheet.Dimension.End.Row)
            {
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
                ExcelWorksheet testExecutionSheet = findSheet(package, "TestExecution");
                getSheetColumnLength(testExecutionSheet);
                getSheetRowLength(testExecutionSheet);

                String[] content = findColumn(testExecutionSheet, "Revision");
                List<String> columnRevision = filterColumn(content);

                ExcelWorksheet testPlanSheet = findSheet(package, "TestPlan");
                getSheetColumnLength(testPlanSheet);
                getSheetRowLength(testPlanSheet);

                int columnIndex = findColumnIndex(testPlanSheet, "Scenario");
                if (testPlanSheet != null && columnIndex > 0) {
                    //Console.WriteLine("Inserting columns...\n");
                    //testPlanSheet.InsertColumn(columnIndex + 1, columnRevision.Count);
                    //Console.WriteLine("Columns inserted!\n");
                }
                package.SaveAs(existingFile);
            }

            Console.WriteLine("The End!\n");
            return Utils.GetFileInfo(fileName).Name;
        }
    }
}
