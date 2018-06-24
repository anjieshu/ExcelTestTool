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

        public static String Run(String fileName)
        {
            Console.WriteLine("Starting running...\nReading file: {0}", fileName);
            Console.WriteLine();

            FileInfo existingFile = new FileInfo(fileName);
            using (ExcelPackage package = new ExcelPackage(existingFile)) {
                ExcelWorksheet testExecutionSheet = findSheet(package, "TestExecution");
                String[] content = findColumn(testExecutionSheet, "Revision");
                List<String> columnRevision = filterColumn(content);

                ExcelWorksheet testPlanSheet = findSheet(package, "TestPlan");
                int columnIndex = findColumnIndex(testPlanSheet, "Scenario");
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

        public static string RunSample1()
        {
            using (var package = new ExcelPackage())
            {
                // Add a new worksheet to the empty workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Inventory");
                //Add the headers
                worksheet.Cells[1, 1].Value = "ID";
                worksheet.Cells[1, 2].Value = "Product";
                worksheet.Cells[1, 3].Value = "Quantity";
                worksheet.Cells[1, 4].Value = "Price";
                worksheet.Cells[1, 5].Value = "Value";

                //Add some items...
                worksheet.Cells["A2"].Value = 12001;
                worksheet.Cells["B2"].Value = "Nails";
                worksheet.Cells["C2"].Value = 37;
                worksheet.Cells["D2"].Value = 3.99;

                worksheet.Cells["A3"].Value = 12002;
                worksheet.Cells["B3"].Value = "Hammer";
                worksheet.Cells["C3"].Value = 5;
                worksheet.Cells["D3"].Value = 12.10;

                worksheet.Cells["A4"].Value = 12003;
                worksheet.Cells["B4"].Value = "Saw";
                worksheet.Cells["C4"].Value = 12;
                worksheet.Cells["D4"].Value = 15.37;

                //Add a formula for the value-column
                worksheet.Cells["E2:E4"].Formula = "C2*D2";

                //Ok now format the values;
                using (var range = worksheet.Cells[1, 1, 1, 5])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.DarkBlue);
                    range.Style.Font.Color.SetColor(Color.White);
                }

                worksheet.Cells["A5:E5"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                worksheet.Cells["A5:E5"].Style.Font.Bold = true;

                worksheet.Cells[5, 3, 5, 5].Formula = string.Format("SUBTOTAL(9,{0})", new ExcelAddress(2, 3, 4, 3).Address);
                worksheet.Cells["C2:C5"].Style.Numberformat.Format = "#,##0";
                worksheet.Cells["D2:E5"].Style.Numberformat.Format = "#,##0.00";

                //Create an autofilter for the range
                worksheet.Cells["A1:E4"].AutoFilter = true;

                worksheet.Cells["A2:A4"].Style.Numberformat.Format = "@";   //Format as text

                //There is actually no need to calculate, Excel will do it for you, but in some cases it might be useful. 
                //For example if you link to this workbook from another workbook or you will open the workbook in a program that hasn't a calculation engine or 
                //you want to use the result of a formula in your program.
                worksheet.Calculate();

                worksheet.Cells.AutoFitColumns(0);  //Autofit columns for all cells

                // lets set the header text 
                worksheet.HeaderFooter.OddHeader.CenteredText = "&24&U&\"Arial,Regular Bold\" Inventory";
                // add the page number to the footer plus the total number of pages
                worksheet.HeaderFooter.OddFooter.RightAlignedText =
                    string.Format("Page {0} of {1}", ExcelHeaderFooter.PageNumber, ExcelHeaderFooter.NumberOfPages);
                // add the sheet name to the footer
                worksheet.HeaderFooter.OddFooter.CenteredText = ExcelHeaderFooter.SheetName;
                // add the file path to the footer
                worksheet.HeaderFooter.OddFooter.LeftAlignedText = ExcelHeaderFooter.FilePath + ExcelHeaderFooter.FileName;

                worksheet.PrinterSettings.RepeatRows = worksheet.Cells["1:2"];
                worksheet.PrinterSettings.RepeatColumns = worksheet.Cells["A:G"];

                // Change the sheet view to show it in page layout mode
                worksheet.View.PageLayoutView = true;

                // set some document properties
                package.Workbook.Properties.Title = "Invertory";
                package.Workbook.Properties.Author = "Jan Källman";
                package.Workbook.Properties.Comments = "This sample demonstrates how to create an Excel 2007 workbook using EPPlus";

                // set some extended property values
                package.Workbook.Properties.Company = "AdventureWorks Inc.";

                // set some custom property values
                package.Workbook.Properties.SetCustomPropertyValue("Checked by", "Jan Källman");
                package.Workbook.Properties.SetCustomPropertyValue("AssemblyName", "EPPlus");

                var xlFile = Utils.GetFileInfo("sample1.xlsx");
                // save our new workbook in the output directory and we are done!
                package.SaveAs(xlFile);
                return xlFile.FullName;
            }
        }
    }
}
