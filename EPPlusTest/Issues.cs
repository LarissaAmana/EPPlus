using System;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Reflection;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Logging;
using OfficeOpenXml.Style;
using System.Data;
using OfficeOpenXml.Table;
using System.Collections.Generic;
using System.Runtime.ExceptionServices;
using EPPlusTest.Properties;
using OfficeOpenXml.Table.PivotTable;
using Rhino.Mocks;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis.TokenSeparatorHandlers;


namespace EPPlusTest
{
    [TestClass]
    public class Issues : ITokenIndexProvider
    {

        private ITokenFactory _tokenFactory;
        private INameValueProvider _nameValueProvider;
        private int _index = 0;


        [TestInitialize]
        public void Initialize()
        {
            if (!Directory.Exists(@"c:\Temp"))
            {
                Directory.CreateDirectory(@"c:\Temp");
            }

            if (!Directory.Exists(@"c:\Temp\bug"))
            {
                Directory.CreateDirectory(@"c:\Temp\bug");
            }
        }

        [TestMethod, Ignore]
        public void Issue15052()
        {
            var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("test");
            ws.Cells["A1:A4"].Value = 1;
            ws.Cells["B1:B4"].Value = 2;

            ws.Cells[1, 1, 4, 1]
                .Style.Numberformat.Format = "#,##0.00;[Red]-#,##0.00";

            ws.Cells[1, 2, 5, 2]
                .Style.Numberformat.Format = "#,##0;[Red]-#,##0";

            p.SaveAs(new FileInfo(@"c:\temp\style.xlsx"));
        }

        [TestMethod]
        public void Issue15041()
        {
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("Test");
                ws.Cells["A1"].Value = 202100083;
                ws.Cells["A1"].Style.Numberformat.Format = "00\\.00\\.00\\.000\\.0";
                Assert.AreEqual("02.02.10.008.3", ws.Cells["A1"].Text);
                ws.Dispose();
            }
        }

        [TestMethod]
        public void Issue15031()
        {
            var d = OfficeOpenXml.Utils.ConvertUtil.GetValueDouble(new TimeSpan(35, 59, 1));
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("Test");
                ws.Cells["A1"].Value = d;
                ws.Cells["A1"].Style.Numberformat.Format = "[t]:mm:ss";
                ws.Dispose();
            }
        }

        [TestMethod]
        public void Issue15022()
        {
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("Test");
                ws.Cells.AutoFitColumns();
                ws.Cells["A1"].Style.Numberformat.Format = "0";
                ws.Cells.AutoFitColumns();
            }
        }

        [TestMethod]
        public void Issue15056()
        {
            var path = @"C:\temp\output.xlsx";
            var file = new FileInfo(path);
            file.Delete();
            using (var ep = new ExcelPackage(file))
            {
                var s = ep.Workbook.Worksheets.Add("test");
                s.Cells["A1:A2"].Formula = ""; // or null, or non-empty whitespace, with same result
                ep.Save();
            }

        }

        [Ignore]
        [TestMethod]
        public void Issue15058()
        {
            System.IO.FileInfo newFile = new System.IO.FileInfo(@"C:\Temp\output.xlsx");
            ExcelPackage excelP = new ExcelPackage(newFile);
            ExcelWorksheet ws = excelP.Workbook.Worksheets[1];
        }

        [Ignore]
        [TestMethod]
        public void Issue15063()
        {
            System.IO.FileInfo newFile = new System.IO.FileInfo(@"C:\Temp\bug\TableFormula.xlsx");
            ExcelPackage excelP = new ExcelPackage(newFile);
            ExcelWorksheet ws = excelP.Workbook.Worksheets[1];
            ws.Calculate();
        }

        [Ignore]
        [TestMethod]
        public void Issue15112()
        {
            System.IO.FileInfo case1 = new System.IO.FileInfo(@"c:\temp\bug\src\src\DeleteRowIssue\Template.xlsx");
            var p = new ExcelPackage(case1);
            var first = p.Workbook.Worksheets[1];
            first.DeleteRow(5);
            p.SaveAs(new System.IO.FileInfo(@"c:\temp\bug\DeleteCol_case1.xlsx"));

            var case2 = new System.IO.FileInfo(@"c:\temp\bug\src2\DeleteRowIssue\Template.xlsx");
            p = new ExcelPackage(case2);
            first = p.Workbook.Worksheets[1];
            first.DeleteRow(5);
            p.SaveAs(new System.IO.FileInfo(@"c:\temp\bug\DeleteCol_case2.xlsx"));
        }

        [Ignore]
        [TestMethod]
        public void Issue15118()
        {
            using (var package = new OfficeOpenXml.ExcelPackage(new FileInfo(@"c:\temp\bugOutput.xlsx"),
                new FileInfo(@"c:\temp\bug\DeleteRowIssue\Template.xlsx")))
            {
                ExcelWorkbook workBook = package.Workbook;
                var worksheet = workBook.Worksheets[1];
                worksheet.Cells[9, 6, 9, 7].Merge = true;
                worksheet.Cells[9, 8].Merge = false;

                worksheet.DeleteRow(5);
                worksheet.DeleteColumn(5);
                worksheet.DeleteColumn(5);
                worksheet.DeleteColumn(5);
                worksheet.DeleteColumn(5);

                package.Save();
            }
        }

        [Ignore]
        [TestMethod]
        public void Issue15109()
        {
            System.IO.FileInfo newFile = new System.IO.FileInfo(@"C:\Temp\bug\test01.xlsx");
            ExcelPackage excelP = new ExcelPackage(newFile);
            ExcelWorksheet ws = excelP.Workbook.Worksheets[1];
            Assert.AreEqual("A1:Z75", ws.Dimension.Address);
            excelP.Dispose();

            newFile = new System.IO.FileInfo(@"C:\Temp\bug\test02.xlsx");
            excelP = new ExcelPackage(newFile);
            ws = excelP.Workbook.Worksheets[1];
            Assert.AreEqual("A1:AF501", ws.Dimension.Address);
            excelP.Dispose();

            newFile = new System.IO.FileInfo(@"C:\Temp\bug\test03.xlsx");
            excelP = new ExcelPackage(newFile);
            ws = excelP.Workbook.Worksheets[1];
            Assert.AreEqual("A1:AD406", ws.Dimension.Address);
            excelP.Dispose();
        }

        [Ignore]
        [TestMethod]
        public void Issue15120()
        {
            var p = new ExcelPackage(new System.IO.FileInfo(@"C:\Temp\bug\pp.xlsx"));
            ExcelWorksheet ws = p.Workbook.Worksheets["tum_liste"];
            ExcelWorksheet wPvt = p.Workbook.Worksheets.Add("pvtSheet");
            var pvSh = wPvt.PivotTables.Add(wPvt.Cells["B5"], ws.Cells[ws.Dimension.Address.ToString()], "pvtS");

            //p.Save();
        }

        [TestMethod]
        public void Issue15113()
        {
            var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("t");
            ws.Cells["A1"].Value = " Performance Update";
            ws.Cells["A1:H1"].Merge = true;
            ws.Cells["A1:H1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
            ws.Cells["A1:H1"].Style.Font.Size = 14;
            ws.Cells["A1:H1"].Style.Font.Color.SetColor(Color.Red);
            ws.Cells["A1:H1"].Style.Font.Bold = true;
            p.SaveAs(new FileInfo(@"c:\temp\merge.xlsx"));
        }

        [TestMethod]
        public void Issue15141()
        {
            using (ExcelPackage package = new ExcelPackage())
            using (ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Test"))
            {
                sheet.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);
                sheet.Cells[1, 1, 1, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                sheet.Cells[1, 5, 2, 5].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                ExcelColumn column = sheet.Column(3); // fails with exception
            }
        }

        [TestMethod, Ignore]
        public void Issue15145()
        {
            using (ExcelPackage p = new ExcelPackage(new System.IO.FileInfo(@"C:\Temp\bug\ColumnInsert.xlsx")))
            {
                ExcelWorksheet ws = p.Workbook.Worksheets[1];
                ws.InsertColumn(12, 3);
                ws.InsertRow(30, 3);
                ws.DeleteRow(31, 1);
                ws.DeleteColumn(7, 1);
                p.SaveAs(new System.IO.FileInfo(@"C:\Temp\bug\InsertCopyFail.xlsx"));
            }
        }

        [TestMethod, Ignore]
        public void Issue15150()
        {
            var template = new FileInfo(@"c:\temp\bug\ClearIssue.xlsx");
            const string output = @"c:\temp\bug\ClearIssueSave.xlsx";

            using (var pck = new ExcelPackage(template, false))
            {
                var ws = pck.Workbook.Worksheets[1];
                ws.Cells["A2:C3"].Value = "Test";
                var c = ws.Cells["B2:B3"];
                c.Clear();

                pck.SaveAs(new FileInfo(output));
            }
        }

        [TestMethod, Ignore]
        public void Issue15146()
        {
            var template = new FileInfo(@"c:\temp\bug\CopyFail.xlsx");
            const string output = @"c:\temp\bug\CopyFail-Save.xlsx";

            using (var pck = new ExcelPackage(template, false))
            {
                var ws = pck.Workbook.Worksheets[3];

                //ws.InsertColumn(3, 1);
                CustomColumnInsert(ws, 3, 1);

                pck.SaveAs(new FileInfo(output));
            }
        }

        private static void CustomColumnInsert(ExcelWorksheet ws, int column, int columns)
        {
            var source = ws.Cells[1, column, ws.Dimension.End.Row, ws.Dimension.End.Column];
            var dest = ws.Cells[1, column + columns, ws.Dimension.End.Row, ws.Dimension.End.Column + columns];
            source.Copy(dest);
        }

        [TestMethod]
        public void Issue15123()
        {
            var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("t");
            using (var dt = new DataTable())
            {
                dt.Columns.Add("String", typeof(string));
                dt.Columns.Add("Int", typeof(int));
                dt.Columns.Add("Bool", typeof(bool));
                dt.Columns.Add("Double", typeof(double));
                dt.Columns.Add("Date", typeof(DateTime));

                var dr = dt.NewRow();
                dr[0] = "Row1";
                dr[1] = 1;
                dr[2] = true;
                dr[3] = 1.5;
                dr[4] = new DateTime(2014, 12, 30);
                dt.Rows.Add(dr);

                dr = dt.NewRow();
                dr[0] = "Row2";
                dr[1] = 2;
                dr[2] = false;
                dr[3] = 2.25;
                dr[4] = new DateTime(2014, 12, 31);
                dt.Rows.Add(dr);

                ws.Cells["A1"].LoadFromDataTable(dt, true);
                ws.Cells["D2:D3"].Style.Numberformat.Format = "(* #,##0.00);_(* (#,##0.00);_(* \"-\"??_);(@)";

                ws.Cells["E2:E3"].Style.Numberformat.Format = "mm/dd/yyyy";
                ws.Cells.AutoFitColumns();
                Assert.AreNotEqual(ws.Cells[2, 5].Text, "");
            }
        }

        [TestMethod]
        public void Issue15128()
        {
            var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("t");
            ws.Cells["A1"].Value = 1;
            ws.Cells["B1"].Value = 2;
            ws.Cells["B2"].Formula = "A1+$B$1";
            ws.Cells["C1"].Value = "Test";
            ws.Cells["A1:B2"].Copy(ws.Cells["C1"]);
            ws.Cells["B2"].Copy(ws.Cells["D1"]);
            p.SaveAs(new FileInfo(@"c:\temp\bug\copy.xlsx"));
        }

        [TestMethod]
        public void IssueMergedCells()
        {
            var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("t");
            ws.Cells["A1:A5,C1:C8"].Merge = true;
            ws.Cells["C1:C8"].Merge = false;
            ws.Cells["A1:A8"].Merge = false;
            p.Dispose();
        }

        [Ignore]
        [TestMethod]
        public void Issue15158()
        {
            using (var package = new OfficeOpenXml.ExcelPackage(new FileInfo(@"c:\temp\Output.xlsx"),
                new FileInfo(@"C:\temp\bug\DeleteColFormula\FormulasIssue\demo.xlsx")))
            {
                ExcelWorkbook workBook = package.Workbook;
                ExcelWorksheet worksheet = workBook.Worksheets[1];

                //string column = ColumnIndexToColumnLetter(28);
                worksheet.DeleteColumn(28);

                if (worksheet.Cells["AA19"].Formula != "")
                {
                    throw new Exception("this cell should not have formula");
                }

                package.Save();
            }
        }

        public class cls1
        {
            public int prop1 { get; set; }
        }

        public class cls2 : cls1
        {
            public string prop2 { get; set; }
        }

        [TestMethod]
        public void LoadFromColIssue()
        {
            var l = new List<cls1>();

            var c2 = new cls2() {prop1 = 1, prop2 = "test1"};
            l.Add(c2);

            var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Test");

            ws.Cells["A1"].LoadFromCollection(l, true, TableStyles.Light16, BindingFlags.Instance | BindingFlags.Public,
                new MemberInfo[] {typeof(cls2).GetProperty("prop2")});
        }

        [TestMethod]
        public void Issue15168()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Test");
                ws.Cells[1, 1].Value = "A1";
                ws.Cells[2, 1].Value = "A2";

                ws.Cells[2, 1].Value = ws.Cells[1, 1].Value;
                Assert.AreEqual("A1", ws.Cells[1, 1].Value);
            }
        }

        [Ignore]
        [TestMethod]
        public void Issue15159()
        {
            var fs = new FileStream(@"C:\temp\bug\DeleteColFormula\FormulasIssue\demo.xlsx", FileMode.OpenOrCreate);
            using (var package = new OfficeOpenXml.ExcelPackage(fs))
            {
                package.Save();
            }

            fs.Seek(0, SeekOrigin.Begin);
            var fs2 = fs;
        }

        [TestMethod]
        public void Issue15179()
        {
            using (var package = new OfficeOpenXml.ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("MergeDeleteBug");
                ws.Cells["E3:F3"].Merge = true;
                ws.Cells["E3:F3"].Merge = false;
                ws.DeleteRow(2, 6);
                ws.Cells["A1"].Value = 0;
                var s = ws.Cells["A1"].Value.ToString();

            }
        }

        [Ignore]
        [TestMethod]
        public void Issue15169()
        {
            FileInfo fileInfo = new FileInfo(@"C:\temp\bug\issue\input.xlsx");

            ExcelPackage excelPackage = new ExcelPackage(fileInfo);
            {
                string sheetName = "Labour Costs";

                ExcelWorksheet ws = excelPackage.Workbook.Worksheets[sheetName];
                excelPackage.Workbook.Worksheets.Delete(ws);

                ws = excelPackage.Workbook.Worksheets.Add(sheetName);

                excelPackage.SaveAs(new FileInfo(@"C:\temp\bug\issue\output2.xlsx"));
            }
        }

        [Ignore]
        [TestMethod]
        public void Issue15172()
        {
            FileInfo fileInfo = new FileInfo(@"C:\temp\bug\book2.xlsx");

            ExcelPackage excelPackage = new ExcelPackage(fileInfo);
            {
                ExcelWorksheet ws = excelPackage.Workbook.Worksheets[1];

                Assert.AreEqual("IF($R10>=X$2,1,0)", ws.Cells["X10"].Formula);
                ws.Calculate();
                Assert.AreEqual(0D, ws.Cells["X10"].Value);
            }
        }

        [Ignore]
        [TestMethod]
        public void Issue15174()
        {
            using (ExcelPackage package = new ExcelPackage(new FileInfo(@"C:\temp\bug\MyTemplate.xlsx")))
            {
                package.Workbook.Worksheets[1].Column(2).Style.Numberformat.Format = "dd/mm/yyyy";

                package.SaveAs(new FileInfo(@"C:\temp\bug\MyTemplate2.xlsx"));
            }
        }

        [Ignore]
        [TestMethod]
        public void PictureIssue()
        {
            var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("t");
            ws.Drawings.AddPicture("Test", new FileInfo(@"c:\temp\bug\2152228.jpg"));
            p.SaveAs(new FileInfo(@"c:\temp\bug\pic.xlsx"));
        }

        [Ignore]
        [TestMethod]
        public void Issue14988()
        {
            var guid = Guid.NewGuid().ToString("N");
            using (var outputStream = new FileStream(@"C:\temp\" + guid + ".xlsx", FileMode.Create))
            {
                using (var inputStream = new FileStream(@"C:\temp\bug2.xlsx", FileMode.Open))
                {
                    using (var package = new ExcelPackage(outputStream, inputStream, "Test"))
                    {
                        var ws = package.Workbook.Worksheets.Add("Test empty");
                        ws.Cells["A1"].Value = "Test";
                        package.Encryption.Password = "Test2";
                        package.Save();
                        //package.SaveAs(new FileInfo(@"c:\temp\test2.xlsx"));
                    }
                }
            }
        }

        [TestMethod, Ignore]
        public void Issue15173_1()
        {
            using (var pck = new ExcelPackage(new FileInfo(@"c:\temp\EPPlusIssues\Excel01.xlsx")))
            {
                var sw = new Stopwatch();
                //pck.Workbook.FormulaParser.Configure(x => x.AttachLogger(LoggerFactory.CreateTextFileLogger(new FileInfo(@"c:\Temp\log1.txt"))));
                sw.Start();
                var ws = pck.Workbook.Worksheets.First();
                pck.Workbook.Calculate();
                Assert.AreEqual("20L2300", ws.Cells["F4"].Value);
                Assert.AreEqual("20K2E01", ws.Cells["F5"].Value);
                var f7Val = pck.Workbook.Worksheets["MODELLO-TIPO PANNELLO"].Cells["F7"].Value;
                Assert.AreEqual(13.445419, Math.Round((double) f7Val, 6));
                sw.Stop();
                Console.WriteLine(sw.Elapsed.TotalSeconds); // approx. 10 seconds

            }
        }

        [TestMethod, Ignore]
        public void Issue15173_2()
        {
            using (var pck = new ExcelPackage(new FileInfo(@"c:\temp\EPPlusIssues\Excel02.xlsx")))
            {
                var sw = new Stopwatch();
                pck.Workbook.FormulaParser.Configure(x =>
                    x.AttachLogger(LoggerFactory.CreateTextFileLogger(new FileInfo(@"c:\Temp\log1.txt"))));
                sw.Start();
                var ws = pck.Workbook.Worksheets.First();
                //ws.Calculate();
                pck.Workbook.Calculate();
                Assert.AreEqual("20L2300", ws.Cells["F4"].Value);
                Assert.AreEqual("20K2E01", ws.Cells["F5"].Value);
                sw.Stop();
                Console.WriteLine(sw.Elapsed.TotalSeconds); // approx. 10 seconds

            }
        }

        [Ignore]
        [TestMethod]
        public void Issue15154()
        {
            Directory.EnumerateFiles(@"c:\temp\bug\ConstructorInvokationNotThreadSafe\").AsParallel().ForAll(file =>
            {
                //lock (_lock)
                //{
                using (var package = new ExcelPackage(new FileStream(file, FileMode.Open)))
                {
                    package.Workbook.Worksheets[1].Cells[1, 1].Value = file;
                    package.SaveAs(new FileInfo(@"c:\temp\bug\ConstructorInvokationNotThreadSafe\new\" +
                                                new FileInfo(file).Name));
                }

                //}
            });

        }

        [Ignore]
        [TestMethod]
        public void Issue15188()
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("test");
                worksheet.Column(6).Style.Numberformat.Format = "mm/dd/yyyy";
                worksheet.Column(7).Style.Numberformat.Format = "mm/dd/yyyy";
                worksheet.Column(8).Style.Numberformat.Format = "mm/dd/yyyy";
                worksheet.Column(10).Style.Numberformat.Format = "mm/dd/yyyy";

                worksheet.Cells[2, 6].Value = DateTime.Today;
                string a = worksheet.Cells[2, 6].Text;
                Assert.AreEqual(DateTime.Today.ToString("MM/dd/yyyy"), a);
            }
        }

        [TestMethod, Ignore]
        public void Issue15194()
        {
            using (var package = new OfficeOpenXml.ExcelPackage(new FileInfo(@"c:\temp\bug\i15194-Save.xlsx"),
                new FileInfo(@"c:\temp\bug\I15194.xlsx")))
            {
                ExcelWorkbook workBook = package.Workbook;
                var worksheet = workBook.Worksheets[1];

                worksheet.Cells["E3:F3"].Merge = false;

                worksheet.DeleteRow(2, 6);

                package.Save();
            }
        }

        [TestMethod, Ignore]
        public void Issue15195()
        {
            using (var package = new OfficeOpenXml.ExcelPackage(new FileInfo(@"c:\temp\bug\i15195_Save.xlsx"),
                new FileInfo(@"c:\temp\bug\i15195.xlsx")))
            {
                ExcelWorkbook workBook = package.Workbook;
                var worksheet = workBook.Worksheets[1];

                worksheet.InsertColumn(8, 2);

                package.Save();
            }
        }

        [TestMethod, Ignore]
        public void Issue14788()
        {
            using (var package = new OfficeOpenXml.ExcelPackage(new FileInfo(@"c:\temp\bug\i15195_Save.xlsx"),
                new FileInfo(@"c:\temp\bug\GetWorkSheetXmlBad.xlsx")))
            {
                ExcelWorkbook workBook = package.Workbook;
                var worksheet = workBook.Worksheets[1];

                worksheet.InsertColumn(8, 2);

                package.Save();
            }
        }

        [TestMethod, Ignore]
        public void Issue15167()
        {
            FileInfo fileInfo = new FileInfo(@"c:\temp\bug\Draw\input.xlsx");

            ExcelPackage excelPackage = new ExcelPackage(fileInfo);
            {
                string sheetName = "Board pack";

                ExcelWorksheet ws = excelPackage.Workbook.Worksheets[sheetName];
                excelPackage.Workbook.Worksheets.Delete(ws);

                ws = excelPackage.Workbook.Worksheets.Add(sheetName);

                excelPackage.SaveAs(new FileInfo(@"c:\temp\bug\output.xlsx"));
            }
        }

        [TestMethod, Ignore]
        public void Issue15198()
        {
            using (var package = new OfficeOpenXml.ExcelPackage(new FileInfo(@"c:\temp\bug\Output.xlsx"),
                new FileInfo(@"c:\temp\bug\demo.xlsx")))
            {
                ExcelWorkbook workBook = package.Workbook;
                var worksheet = workBook.Worksheets[1];

                worksheet.DeleteRow(12);

                package.Save();
            }
        }

        [TestMethod, Ignore]
        public void Issue13492()
        {
            using (var package = new OfficeOpenXml.ExcelPackage(new FileInfo(@"c:\temp\bug\Bug13492.xlsx")))
            {
                ExcelWorkbook workBook = package.Workbook;
                var worksheet = workBook.Worksheets[1];

                var rt = worksheet.Cells["K31"].RichText.Text;

                package.Save();
            }
        }

        [TestMethod, Ignore]
        public void Issue14966()
        {
            using (var package = new ExcelPackage(new FileInfo(@"c:\temp\bug\ssis\FileFromReportingServer2012.xlsx")))
                package.SaveAs(new FileInfo(@"c:\temp\bug\ssis\Corrupted.xlsx"));
        }

        [TestMethod, Ignore]
        public void Issue15200()
        {
            File.Copy(@"C:\temp\bug\EPPlusRangeCopyTest\EPPlusRangeCopyTest\input.xlsx",
                @"C:\temp\bug\EPPlusRangeCopyTest\EPPlusRangeCopyTest\output.xlsx", true);

            using (var p =
                new ExcelPackage(new FileInfo(@"C:\temp\bug\EPPlusRangeCopyTest\EPPlusRangeCopyTest\output.xlsx")))
            {
                var sheet = p.Workbook.Worksheets.First();

                var sourceRange = sheet.Cells[1, 1, 1, 2];
                var resultRange = sheet.Cells[3, 1, 3, 2];
                sourceRange.Copy(resultRange);

                sourceRange = sheet.Cells[1, 1, 1, 7];
                resultRange = sheet.Cells[5, 1, 5, 7];
                sourceRange.Copy(
                    resultRange); // This throws System.ArgumentException: Can't merge and already merged range

                sourceRange = sheet.Cells[1, 1, 1, 7];
                resultRange = sheet.Cells[7, 3, 7, 7];
                sourceRange.Copy(
                    resultRange); // This throws System.ArgumentException: Can't merge and already merged range

                p.Save();
            }
        }

        [TestMethod]
        public void Issue15212()
        {
            var s = "_(\"R$ \"* #,##0.00_);_(\"R$ \"* (#,##0.00);_(\"R$ \"* \"-\"??_);_(@_) )";
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("StyleBug");
                ws.Cells["A1"].Value = 5698633.64;
                ws.Cells["A1"].Style.Numberformat.Format = s;
                var t = ws.Cells["A1"].Text;
            }
        }

        [TestMethod]
        public void Issue_KeyNotFoundSaveProblem()
        {
            using (var file = new MemoryStream())
            {
                file.Write(EPPlusTest.Properties.Resources.Issue_KeyNotFoundSaveProblem, 0,
                    EPPlusTest.Properties.Resources.Issue_KeyNotFoundSaveProblem.Length);
                using (ExcelPackage package = new ExcelPackage(file))

                {

                    package.SaveAs(new FileInfo(@"Pleas insert a path"));
                }
            }
        }

        [TestMethod, Ignore]
        public void Issue15213()
        {
            using (var p = new ExcelPackage(new FileInfo(@"c:\temp\bug\ExcelClearDemo\exceltestfile.xlsx")))
            {
                foreach (var ws in p.Workbook.Worksheets)
                {
                    ws.Cells[1023, 1, ws.Dimension.End.Row - 2, ws.Dimension.End.Column].Clear();
                    Assert.AreNotEqual(ws.Dimension, null);
                }

                foreach (var cell in p.Workbook.Worksheets[2].Cells)
                {
                    Console.WriteLine(cell);
                }

                p.SaveAs(new FileInfo(@"c:\temp\bug\ExcelClearDemo\exceltestfile-save.xlsx"));
            }
        }

        [TestMethod, Ignore]
        public void Issuer15217()
        {

            using (var p = new ExcelPackage(new FileInfo(@"c:\temp\bug\FormatRowCol.xlsx")))
            {
                var ws = p.Workbook.Worksheets.Add("fmt");
                ws.Row(1).Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Row(1).Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                ws.Cells["A1:B2"].Value = 1;
                ws.Column(1).Style.Numberformat.Format = "yyyy-mm-dd hh:mm";
                ws.Column(2).Style.Numberformat.Format = "#,##0";
                p.Save();
            }
        }

        [TestMethod, Ignore]
        public void Issuer15228()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("colBug");

                var col = ws.Column(7);
                col.ColumnMax = 8;
                col.Hidden = true;

                var col8 = ws.Column(8);
                Assert.AreEqual(true, col8.Hidden);
            }
        }

        [TestMethod, Ignore]
        public void Issue15234()
        {
            using (var p = new ExcelPackage(new FileInfo(@"c:\temp\bug\merge2\input.xlsx")))
            {
                var sheet = p.Workbook.Worksheets.First();

                var sourceRange = sheet.Cells["1:4"];

                sheet.InsertRow(5, 4);

                var resultRange = sheet.Cells["5:8"];
                sourceRange.Copy(resultRange);

                p.Save();
            }
        }

        [TestMethod]
        /**** Pivottable issue ****/
        public void Issue()
        {
            DirectoryInfo outputDir = new DirectoryInfo(@"c:\ExcelPivotTest");
            FileInfo MyFile = new FileInfo(@"c:\temp\bug\pivottable.xlsx");
            LoadData(MyFile);
            BuildPivotTable1(MyFile);
            BuildPivotTable2(MyFile);
        }

        private void LoadData(FileInfo MyFile)
        {
            if (MyFile.Exists)
            {
                MyFile.Delete(); // ensures we create a new workbook
            }

            using (ExcelPackage EP = new ExcelPackage(MyFile))
            {
                // add a new worksheet to the empty workbook
                ExcelWorksheet wsData = EP.Workbook.Worksheets.Add("Data");
                //Add the headers
                wsData.Cells[1, 1].Value = "INVOICE_DATE";
                wsData.Cells[1, 2].Value = "TOTAL_INVOICE_PRICE";
                wsData.Cells[1, 3].Value = "EXTENDED_PRICE_VARIANCE";
                wsData.Cells[1, 4].Value = "AUDIT_LINE_STATUS";
                wsData.Cells[1, 5].Value = "RESOLUTION_STATUS";
                wsData.Cells[1, 6].Value = "COUNT";

                //Add some items...
                wsData.Cells["A2"].Value = Convert.ToDateTime("04/2/2012");
                wsData.Cells["B2"].Value = 33.63;
                wsData.Cells["C2"].Value = (-.87);
                wsData.Cells["D2"].Value = "Unfavorable Price Variance";
                wsData.Cells["E2"].Value = "Pending";
                wsData.Cells["F2"].Value = 1;

                wsData.Cells["A3"].Value = Convert.ToDateTime("04/2/2012");
                wsData.Cells["B3"].Value = 43.14;
                wsData.Cells["C3"].Value = (-1.29);
                wsData.Cells["D3"].Value = "Unfavorable Price Variance";
                wsData.Cells["E3"].Value = "Pending";
                wsData.Cells["F3"].Value = 1;

                wsData.Cells["A4"].Value = Convert.ToDateTime("11/8/2011");
                wsData.Cells["B4"].Value = 55;
                wsData.Cells["C4"].Value = (-2.87);
                wsData.Cells["D4"].Value = "Unfavorable Price Variance";
                wsData.Cells["E4"].Value = "Pending";
                wsData.Cells["F4"].Value = 1;

                wsData.Cells["A5"].Value = Convert.ToDateTime("11/8/2011");
                wsData.Cells["B5"].Value = 38.72;
                wsData.Cells["C5"].Value = (-5.00);
                wsData.Cells["D5"].Value = "Unfavorable Price Variance";
                wsData.Cells["E5"].Value = "Pending";
                wsData.Cells["F5"].Value = 1;

                wsData.Cells["A6"].Value = Convert.ToDateTime("3/4/2011");
                wsData.Cells["B6"].Value = 77.44;
                wsData.Cells["C6"].Value = (-1.55);
                wsData.Cells["D6"].Value = "Unfavorable Price Variance";
                wsData.Cells["E6"].Value = "Pending";
                wsData.Cells["F6"].Value = 1;

                wsData.Cells["A7"].Value = Convert.ToDateTime("3/4/2011");
                wsData.Cells["B7"].Value = 127.55;
                wsData.Cells["C7"].Value = (-10.50);
                wsData.Cells["D7"].Value = "Unfavorable Price Variance";
                wsData.Cells["E7"].Value = "Pending";
                wsData.Cells["F7"].Value = 1;

                using (var range = wsData.Cells[2, 1, 7, 1])
                {
                    range.Style.Numberformat.Format = "mm-dd-yy";
                }

                wsData.Cells.AutoFitColumns(0);
                EP.Save();
            }
        }

        private void BuildPivotTable1(FileInfo MyFile)
        {
            using (ExcelPackage ep = new ExcelPackage(MyFile))
            {

                var wsData = ep.Workbook.Worksheets["Data"];
                var totalRows = wsData.Dimension.Address;
                ExcelRange data = wsData.Cells[totalRows];

                var wsAuditPivot = ep.Workbook.Worksheets.Add("Pivot1");

                var pivotTable1 = wsAuditPivot.PivotTables.Add(wsAuditPivot.Cells["A7:C30"], data, "PivotAudit1");
                pivotTable1.ColumGrandTotals = true;
                var rowField = pivotTable1.RowFields.Add(pivotTable1.Fields["INVOICE_DATE"]);


                rowField.AddDateGrouping(eDateGroupBy.Years);
                var yearField = pivotTable1.Fields.GetDateGroupField(eDateGroupBy.Years);
                yearField.Name = "Year";

                var rowField2 = pivotTable1.RowFields.Add(pivotTable1.Fields["AUDIT_LINE_STATUS"]);

                var TotalSpend = pivotTable1.DataFields.Add(pivotTable1.Fields["TOTAL_INVOICE_PRICE"]);
                TotalSpend.Name = "Total Spend";
                TotalSpend.Format = "$##,##0";


                var CountInvoicePrice = pivotTable1.DataFields.Add(pivotTable1.Fields["COUNT"]);
                CountInvoicePrice.Name = "Total Lines";
                CountInvoicePrice.Format = "##,##0";

                pivotTable1.DataOnRows = false;
                ep.Save();
                ep.Dispose();

            }

        }

        private void BuildPivotTable2(FileInfo MyFile)
        {
            using (ExcelPackage ep = new ExcelPackage(MyFile))
            {

                var wsData = ep.Workbook.Worksheets["Data"];
                var totalRows = wsData.Dimension.Address;
                ExcelRange data = wsData.Cells[totalRows];

                var wsAuditPivot = ep.Workbook.Worksheets.Add("Pivot2");

                var pivotTable1 = wsAuditPivot.PivotTables.Add(wsAuditPivot.Cells["A7:C30"], data, "PivotAudit2");
                pivotTable1.ColumGrandTotals = true;
                var rowField = pivotTable1.RowFields.Add(pivotTable1.Fields["INVOICE_DATE"]);


                rowField.AddDateGrouping(eDateGroupBy.Years);
                var yearField = pivotTable1.Fields.GetDateGroupField(eDateGroupBy.Years);
                yearField.Name = "Year";

                var rowField2 = pivotTable1.RowFields.Add(pivotTable1.Fields["AUDIT_LINE_STATUS"]);

                var TotalSpend = pivotTable1.DataFields.Add(pivotTable1.Fields["TOTAL_INVOICE_PRICE"]);
                TotalSpend.Name = "Total Spend";
                TotalSpend.Format = "$##,##0";


                var CountInvoicePrice = pivotTable1.DataFields.Add(pivotTable1.Fields["COUNT"]);
                CountInvoicePrice.Name = "Total Lines";
                CountInvoicePrice.Format = "##,##0";

                pivotTable1.DataOnRows = false;
                ep.Save();
                ep.Dispose();

            }

        }

        [TestMethod, Ignore]
        public void issue15249()
        {
            using (var exfile = new ExcelPackage(new FileInfo(@"c:\temp\bug\Boldtextcopy.xlsx")))
            {
                exfile.Workbook.Worksheets.Copy("sheet1", "copiedSheet");
                exfile.SaveAs(new FileInfo(@"c:\temp\bug\Boldtextcopy2.xlsx"));
            }
        }

        [TestMethod, Ignore]
        public void issue15300()
        {
            using (var exfile = new ExcelPackage(new FileInfo(@"c:\temp\bug\headfootpic.xlsx")))
            {
                exfile.Workbook.Worksheets.Copy("sheet1", "copiedSheet");
                exfile.SaveAs(new FileInfo(@"c:\temp\bug\headfootpic_save.xlsx"));
            }

        }

        [TestMethod, Ignore]
        public void issue15295()
        {
            using (var exfile = new ExcelPackage(new FileInfo(@"C:\temp\bug\pivot issue\input.xlsx")))
            {
                exfile.SaveAs(new FileInfo(@"C:\temp\bug\pivot issue\pivotcoldup.xlsx"));
            }

        }

        [TestMethod, Ignore]
        public void issue15282()
        {
            using (var exfile = new ExcelPackage(new FileInfo(@"C:\temp\bug\pivottable-table.xlsx")))
            {
                exfile.SaveAs(new FileInfo(@"C:\temp\bug\pivot issue\pivottab-tab-save.xlsx"));
            }

        }

        [TestMethod, Ignore]
        public void Issues14699()
        {
            FileInfo newFile = new FileInfo(string.Format("c:\\temp\\bug\\EPPlus_Issue14699.xlsx",
                System.IO.Directory.GetCurrentDirectory()));
            OfficeOpenXml.ExcelPackage pkg = new ExcelPackage(newFile);
            ExcelWorksheet wksheet = pkg.Workbook.Worksheets.Add("Issue14699");
            // Initialize a small range
            for (int row = 1; row < 11; row++)
            {
                for (int col = 1; col < 11; col++)
                {
                    wksheet.Cells[row, col].Value = string.Format("{0}{1}", "ABCDEFGHIJKLMNOPQRSTUVWXYZ"[col - 1], row);
                }
            }

            wksheet.View.FreezePanes(3, 3);
            pkg.Save();

        }

        [TestMethod, Ignore]
        public void Issue15382()
        {
            using (var exfile = new ExcelPackage(new FileInfo(@"c:\temp\bug\Text Run Issue.xlsx")))
            {
                exfile.SaveAs(new FileInfo(@"C:\temp\bug\inlinText.xlsx"));
            }
        }

        [TestMethod, Ignore]
        public void Issue15380()
        {
            using (var exfile = new ExcelPackage(new FileInfo(@"c:\temp\bug\dotinname.xlsx")))
            {
                var v = exfile.Workbook.Worksheets["sheet1.3"].Names["Test.Name"].Value;
                Assert.AreEqual(v, 1);
            }
        }

        [TestMethod, Ignore]
        public void Issue15378()
        {
            using (var p = new ExcelPackage(new FileInfo(@"c:\temp\bubble.xlsx")))
            {
                var c = p.Workbook.Worksheets[1].Drawings[0] as ExcelBubbleChart;
                var cs = c.Series[0] as ExcelBubbleChartSerie;
            }
        }

        [TestMethod]
        public void Issue15377()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("ws1");
                ws.Cells["A1"].Value = (double?) 1;
                var v = ws.GetValue<double?>(1, 1);
            }
        }

        [TestMethod]
        public void Issue15374()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("RT");
                var r = ws.Cells["A1"];
                r.RichText.Text = "Cell 1";
                r["A2"].RichText.Add("Cell 2");
                p.SaveAs(new FileInfo(@"c:\temp\rt.xlsx"));
            }
        }

        [TestMethod]
        public void IssueTranslate()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Trans");
                ws.Cells["A1:A2"].Formula = "IF(1=1, \"A's B C\",\"D\") ";
                var fr = ws.Cells["A1:A2"].FormulaR1C1;
                ws.Cells["A1:A2"].FormulaR1C1 = fr;
                Assert.AreEqual("IF(1=1,\"A's B C\",\"D\")", ws.Cells["A2"].Formula);
            }
        }

        [TestMethod]
        public void Issue15397()
        {
            using (var p = new ExcelPackage())
            {
                var workSheet = p.Workbook.Worksheets.Add("styleerror");
                workSheet.Cells["F:G"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells["F:G"].Style.Fill.BackgroundColor.SetColor(Color.Red);

                workSheet.Cells["A:A,C:C"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells["A:A,C:C"].Style.Fill.BackgroundColor.SetColor(Color.Red);

                //And then: 

                workSheet.Cells["A:H"].Style.Font.Color.SetColor(Color.Blue);

                workSheet.Cells["I:I"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                workSheet.Cells["I:I"].Style.Fill.BackgroundColor.SetColor(Color.Red);
                workSheet.Cells["I2"].Style.Fill.BackgroundColor.SetColor(Color.Green);
                workSheet.Cells["I4"].Style.Fill.BackgroundColor.SetColor(Color.Blue);
                workSheet.Cells["I9"].Style.Fill.BackgroundColor.SetColor(Color.Pink);

                workSheet.InsertColumn(2, 2, 9);
                workSheet.Column(45).Width = 0;

                p.SaveAs(new FileInfo(@"c:\temp\styleerror.xlsx"));
            }
        }

        [TestMethod]
        public void Issuer14801()
        {
            using (var p = new ExcelPackage())
            {
                var workSheet = p.Workbook.Worksheets.Add("rterror");
                var cell = workSheet.Cells["A1"];
                cell.RichText.Add("toto: ");
                cell.RichText[0].PreserveSpace = true;
                cell.RichText[0].Bold = true;
                cell.RichText.Add("tata");
                cell.RichText[1].Bold = false;
                cell.RichText[1].Color = Color.Green;
                p.SaveAs(new FileInfo(@"c:\temp\rtpreserve.xlsx"));
            }
        }

        [TestMethod]
        public void Issuer15445()
        {
            using (var p = new ExcelPackage())
            {
                var ws1 = p.Workbook.Worksheets.Add("ws1");
                var ws2 = p.Workbook.Worksheets.Add("ws2");
                ws2.View.SelectedRange = "A1:B3 D12:D15";
                ws2.View.ActiveCell = "D15";
                p.SaveAs(new FileInfo(@"c:\temp\activeCell.xlsx"));
            }
        }

        [TestMethod, Ignore]
        public void Issue15429()
        {
            FileInfo file = new FileInfo(@"c:\temp\original.xlsx");
            using (ExcelPackage excelPackage = new ExcelPackage(file))
            {
                var worksheet = excelPackage.Workbook.Worksheets.Add("Sheet 1");
                var equalsRule = worksheet.ConditionalFormatting.AddEqual(new ExcelAddress(2, 3, 6, 3));
                equalsRule.Formula = "0";
                equalsRule.Style.Fill.BackgroundColor.Color = Color.Blue;
                worksheet.ConditionalFormatting.AddDatabar(new ExcelAddress(4, 4, 4, 4), Color.Red);
                excelPackage.Save();
            }

            using (ExcelPackage excelPackage = new ExcelPackage(file))
            {
                var worksheet = excelPackage.Workbook.Worksheets["Sheet 1"];
                int i = 0;
                foreach (var conditionalFormat in worksheet.ConditionalFormatting)
                {
                    conditionalFormat.Address = new ExcelAddress(5 + i++, 5, 6, 6);
                }

                excelPackage.SaveAs(new FileInfo(@"c:\temp\error.xlsx"));
            }
        }

        [TestMethod, Ignore]
        public void Issue15436()
        {
            FileInfo file = new FileInfo(@"c:\temp\incorrect value.xlsx");
            using (ExcelPackage excelPackage = new ExcelPackage(file))
            {
                Assert.AreEqual(excelPackage.Workbook.Worksheets[1].Cells["A1"].Value, 19120072);
            }
        }

        [TestMethod, Ignore]
        public void Issue13128()
        {
            FileInfo file = new FileInfo(@"c:\temp\students.xlsx");
            using (ExcelPackage excelPackage = new ExcelPackage(file))
            {
                Assert.AreNotEqual(((ExcelChart) excelPackage.Workbook.Worksheets[1].Drawings[0]).Series[0].XSeries,
                    null);
            }
        }

        [TestMethod, Ignore]
        public void Issue15252()
        {
            using (var p = new ExcelPackage())
            {
                var path1 = @"c:\temp\saveerror1.xlsx";
                var path2 = @"c:\temp\saveerror2.xlsx";
                var workSheet = p.Workbook.Worksheets.Add("saveerror");
                workSheet.Cells["A1"].Value = "test";

                // double save OK?
                p.SaveAs(new FileInfo(path1));
                p.SaveAs(new FileInfo(path2));

                // files are identical?
                var md5 = new System.Security.Cryptography.MD5CryptoServiceProvider();
                using (var fs1 = new FileStream(path1, FileMode.Open))
                using (var fs2 = new FileStream(path2, FileMode.Open))
                {
                    var hash1 = String.Join("", md5.ComputeHash(fs1).Select((x) => { return x.ToString(); }));
                    var hash2 = String.Join("", md5.ComputeHash(fs2).Select((x) => { return x.ToString(); }));
                    Assert.AreEqual(hash1, hash2);
                }
            }
        }

        [TestMethod, Ignore]
        public void Issue15469()
        {
            ExcelPackage excelPackage = new ExcelPackage(new FileInfo(@"c:\temp\bug\EPPlus-Bug.xlsx"), true);
            using (FileStream fs = new FileStream(@"c:\temp\bug\EPPlus-Bug-new.xlsx", FileMode.Create))
            {
                excelPackage.SaveAs(fs);
            }
        }

        [TestMethod]
        public void Issue15438()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Test");
                var c = ws.Cells["A1"].Style.Font.Color;
                c.Indexed = 3;
                Assert.AreEqual(c.LookupColor(c), "#FF00FF00");
            }
        }

        [TestMethod, Ignore]
        public void Issue15097()
        {
            using (var pkg = new ExcelPackage())
            {
                var templateFile = ReadTemplateFile(@"c:\temp\bug\test_vorlage3.xlsx");
                using (var ms = new System.IO.MemoryStream(templateFile))
                {
                    using (var tempPkg = new ExcelPackage(ms))
                    {
                        tempPkg.Workbook.Worksheets.Copy(tempPkg.Workbook.Worksheets.First().Name, "Demo");
                    }
                }
            }
        }

        [TestMethod]
        public void Issue15485()
        {
            using (var pkg = new ExcelPackage(new FileInfo(@"c:\temp\bug\PivotChartSeriesIssue.xlsx")))
            {
                var ws = pkg.Workbook.Worksheets[1];
                ws.InsertRow(1, 1);
                ws.InsertColumn(1, 1);
                pkg.Save();
            }
        }

        public static byte[] ReadTemplateFile(string templateName)
        {
            byte[] templateFIle;
            using (System.IO.MemoryStream ms = new System.IO.MemoryStream())
            {
                using (var sw = new System.IO.FileStream(templateName, System.IO.FileMode.Open,
                    System.IO.FileAccess.Read, System.IO.FileShare.ReadWrite))
                {
                    byte[] buffer = new byte[2048];
                    int bytesRead;
                    while ((bytesRead = sw.Read(buffer, 0, buffer.Length)) > 0)
                    {
                        ms.Write(buffer, 0, bytesRead);
                    }
                }

                ms.Position = 0;
                templateFIle = ms.ToArray();
            }

            return templateFIle;

        }

        [TestMethod]
        public void Issue15455()
        {
            using (var pck = new ExcelPackage())
            {

                var sheet1 = pck.Workbook.Worksheets.Add("sheet1");
                var sheet2 = pck.Workbook.Worksheets.Add("Sheet2");
                sheet1.Cells["C2"].Value = 3;
                sheet1.Cells["C3"].Formula = "VLOOKUP(E1, Sheet2!A1:D6, C2, 0)";
                sheet1.Cells["E1"].Value = "d";

                sheet2.Cells["A1"].Value = "d";
                sheet2.Cells["C1"].Value = "dg";
                pck.Workbook.Calculate();
                var c3 = sheet1.Cells["C3"].Value;
                Assert.AreEqual("dg", c3);
            }
        }

        [TestMethod]
        public void Issue15460WithString()
        {
            FileInfo file = new FileInfo("report.xlsx");
            try
            {
                if (file.Exists)
                    file.Delete();
                using (ExcelPackage package = new ExcelPackage(file))
                {
                    var sheet = package.Workbook.Worksheets.Add("New Sheet");
                    sheet.Cells[3, 3].Value = new[] {"value1", "value2", "value3"};
                    package.Save();
                }

                using (ExcelPackage package = new ExcelPackage(file))
                {
                    var sheet = package.Workbook.Worksheets["New Sheet"];
                    Assert.AreEqual("value1", sheet.Cells[3, 3].Value);
                }
            }
            finally
            {
                if (file.Exists)
                    file.Delete();
            }
        }

        [TestMethod]
        public void Issue15460WithNull()
        {
            FileInfo file = new FileInfo("report.xlsx");
            try
            {
                if (file.Exists)
                    file.Delete();
                using (ExcelPackage package = new ExcelPackage(file))
                {
                    var sheet = package.Workbook.Worksheets.Add("New Sheet");
                    sheet.Cells[3, 3].Value = new[] {null, "value2", "value3"};
                    package.Save();
                }

                using (ExcelPackage package = new ExcelPackage(file))
                {
                    var sheet = package.Workbook.Worksheets["New Sheet"];
                    Assert.AreEqual(string.Empty, sheet.Cells[3, 3].Value);
                }
            }
            finally
            {
                if (file.Exists)
                    file.Delete();
            }
        }

        [TestMethod]
        public void Issue15460WithNonStringPrimitive()
        {
            FileInfo file = new FileInfo("report.xlsx");
            try
            {
                if (file.Exists)
                    file.Delete();
                using (ExcelPackage package = new ExcelPackage(file))
                {
                    var sheet = package.Workbook.Worksheets.Add("New Sheet");
                    sheet.Cells[3, 3].Value = new[] {5, 6, 7};
                    package.Save();
                }

                using (ExcelPackage package = new ExcelPackage(file))
                {
                    var sheet = package.Workbook.Worksheets["New Sheet"];
                    Assert.AreEqual((double) 5, sheet.Cells[3, 3].Value);
                }
            }
            finally
            {
                if (file.Exists)
                    file.Delete();
            }
        }

        [TestMethod]
        public void MergeIssue()
        {
            var worksheetPath = Path.Combine(Path.GetTempPath(), @"EPPlus worksheets");
            FileInfo fi = new FileInfo(Path.Combine(worksheetPath, "Example.xlsx"));
            fi.Delete();
            using (ExcelPackage pckg = new ExcelPackage(fi))
            {
                var ws = pckg.Workbook.Worksheets.Add("Example");
                ws.Cells[1, 1, 1, 3].Merge = true;
                ws.Cells[1, 1, 1, 3].Merge = true;
                pckg.Save();
            }
        }



        [TestMethod]
        public void Issue15221()
        {

            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");

                var comment = sheet.Cells["A1"].AddComment("Text", "Author");
                comment.Visible = true;
                Assert.IsTrue(comment.Style.Contains("visible"));
            }
        }


        [TestMethod]
        public void Issue15345()
        {

            string address1 = @"'C:\Projects\epplusamana_new\EPPlusTest\Workbooks\[FormulaTest.xlsx]Sheet1'#REF!C1";
            string address2 =
                @"'C:\Projects\epplusamana_new\EPPlusTest\Workbooks\[FormulaTest.xlsx]Sheet1'SUMME(#REF!A1:B1)";
            Assert.AreEqual(ExcelAddressBase.IsValid(address1), ExcelAddressBase.AddressType.Invalid);
            Assert.AreEqual(ExcelAddressBase.IsValid(address2), ExcelAddressBase.AddressType.Invalid);
        }

        [TestMethod]
        public void Issue15399()
        {
            using (var getReference = new MemoryStream())
            {

                getReference.Write(EPPlusTest.Properties.Resources.TestFile_Issue15399, 0,
                    EPPlusTest.Properties.Resources.TestFile_Issue15399.Length);
                using (ExcelPackage packageGetReference = new ExcelPackage(getReference))
                {

                    Assert.AreEqual(@"[1]Tabelle1!$A$1", packageGetReference.Workbook.Names[0].Formula);
                    Assert.IsFalse(packageGetReference.Workbook.Names[0].Address.Contains(@"file:"));
                }

            }
        }


        [TestMethod]
        public void Issue15361()
        {
            using (var file = new MemoryStream())
            {
                file.Write(EPPlusTest.Properties.Resources.Reproduce_StandardStyle_Issue15361, 0,
                    EPPlusTest.Properties.Resources.Reproduce_StandardStyle_Issue15361.Length);
                using (ExcelPackage package = new ExcelPackage(file))
                {
                    package.Save();
                }

                file.Position = 0;
                using (ExcelPackage package = new ExcelPackage(file))
                {
                    OfficeOpenXml.Style.XmlAccess.ExcelNamedStyleXml testStyle = null;
                    Assert.IsFalse(package.Workbook.Styles.NamedStyles.FindByID("Normal", ref testStyle),
                        "A normalstyle should be found now.");
                    Assert.IsTrue(package.Workbook.Styles.NamedStyles.FindByID("Standard", ref testStyle),
                        "The standardstyle should now be the normal Style");
                    Assert.AreEqual(0, testStyle.StyleXfId,
                        "The Standardstyle should have the same index as the Normalstyle before.");
                    Assert.AreEqual(0, testStyle.Style.Font.Index,
                        "The Standardstyle should have the same index as the Normalstyle before.");
                    Assert.AreEqual(0, testStyle.Style.Numberformat.Index,
                        "The Standardstyle should have the same index as the Normalstyle before.");
                }
            }

        }


        [TestMethod]
        public void RemoveCommentsIssue()
        {
            using (var p =
                new ExcelPackage(new FileInfo(@"C:\Users\larissa.hohaus\Documents\EPPlus\empty_Dokument.xlsx")))
            {
                var sheet = p.Workbook.Worksheets.First();


                var cellB3 = sheet.Cells["B3"];
                var cellB4 = sheet.Cells["B4"];
                var cellB5 = sheet.Cells["B5"];

                cellB3.AddComment("myCommentB3", "me");
                cellB4.AddComment("myCommentB4", "me");
                cellB5.AddComment("myCommentB5", "me");

                foreach (var cell in sheet.Cells)
                {
                    if (cell.Comment != null)
                    {
                        sheet.Comments.Remove(cell.Comment);
                    }
                }

                Assert.IsNull(cellB3.Comment);
                Assert.IsNull(cellB4.Comment);
                Assert.IsNull(cellB5.Comment);


            }
        }


        [TestMethod]
        public void ParserProblem2()
        {
            using (var p =
                new ExcelPackage(
                    new FileInfo(@"C:\Users\larissa.hohaus\Desktop\BugFixTests\SN_T_1506337027_Kapitel.xlsx")))
            {
                var sheet = p.Workbook.Worksheets["Tabelle1"];
                //sheet.NameSpaceManager.LookupNamespace("outerea");

                Dictionary<int, String> _vsheetFormula = new Dictionary<int, string>();
                Dictionary<int, String> _vsheetFormulaExpected = new Dictionary<int, string>();
                _vsheetFormulaExpected[0] = "";
                _vsheetFormulaExpected[1] = "";
                _vsheetFormulaExpected[2] = "";
                _vsheetFormulaExpected[3] = "SUM(B1:B2,B3)";


                //ExcelRangeBase _columnRange = sheet.Cells[1, 3, ExcelPackage.MaxRows, 3];
                ExcelRangeBase _columnRange = sheet.Cells["C:C"];
                String colRange = String.Empty;
                _vsheetFormula = (from r in _columnRange select r).ToDictionary(r => r.Start.Row, r => r.Formula);
                //foreach (var cell in _columnRange)
                //{
                //    //_vsheetFormula[cell.Start.Row] = cell.Formula;
                //    colRange = $"{colRange};{cell.Address}";
                //}


                Assert.AreEqual(_vsheetFormulaExpected, _vsheetFormula);
                Assert.AreEqual(";C1;C2;C3;C4", colRange);

            }
        }

        [TestMethod]
        public void ParserProblem()
        {
            using (var p = new ExcelPackage())
            {
                p.Workbook.Worksheets.Add("first");

                var sheet = p.Workbook.Worksheets.First();

                //var cellA1 = sheet.Cells["A1"];
                var cellA2 = sheet.Cells["A2"];
                var cellA3 = sheet.Cells["A3"];
                var cellA4 = sheet.Cells["A4"];
                var cellA5 = sheet.Cells["A5"];
                var cellA6 = sheet.Cells["A6"];
                var cellA7 = sheet.Cells["A7"];
                var cellA8 = sheet.Cells["A8"];
                //var cellA9 = sheet.Cells["A9"];
                var cellA10 = sheet.Cells["A10"];
                var cellA11 = sheet.Cells["A11"];
                var cellA12 = sheet.Cells["A12"];
                //cellA1.Value = 1;
                cellA2.Value = 1;
                cellA3.Value = 1;
                cellA4.Value = 1;
                cellA5.Value = 1;
                cellA6.Value = 1;
                cellA7.Value = 1;
                cellA8.Value = 1;
                //cellA9.Value = 1;
                cellA10.Value = 1;
                cellA11.Value = 1;
                cellA12.Formula = "SUM(A1:A3,A5,A6,A7,A8,A10,A9,A11)";


                int counterColRange = 0;
                var colRange = String.Empty;
                var Formula = String.Empty;
                var Value = String.Empty;
                var Adress = String.Empty;
                var result = String.Empty;
                var rowId = String.Empty;
                var rowIds = String.Empty;

                foreach (var cell in sheet.Cells["A:A"])
                {
                    counterColRange++;
                    colRange = $"{colRange};{cell.Address}";
                    Formula = cell.Formula;
                    Value = $"{cell.Value}";
                    Adress = cell.Address;
                    rowId = $"{cell.Start.Row}";
                    result = $"{result},{Adress}:{Formula}:{Value}:{rowId}";
                    rowIds = $"{rowIds};{cell.Start.Row}";

                }

                //Assert.AreEqual(";A1;A2;A3;A4;A5;A6;A7;A8;A9;A10;A11;A12", colRange);
                Assert.AreEqual(";A2;A3;A4;A5;A6;A7;A8;A10;A11;A12", colRange);
                //Assert.AreEqual(";1;2;3;4;5;6;7;8;9;10;11;12", rowIds);
                Assert.AreEqual(";2;3;4;5;6;7;8;10;11;12", rowIds);

                //Assert.AreEqual(12, counterColRange);
                Assert.AreEqual(10, counterColRange);


                counterColRange = 0;
                colRange = String.Empty;
                Formula = String.Empty;
                Value = String.Empty;
                Adress = String.Empty;
                result = String.Empty;
                rowId = String.Empty;
                rowIds = String.Empty;
                foreach (var cell in sheet.Cells["A:A"])
                {
                    counterColRange++;
                    colRange = $"{colRange};{cell.Address}";
                    Formula = cell.Formula;
                    Value = $"{cell.Value}";
                    Adress = cell.Address;
                    rowId = $"{cell.Start.Row}";
                    result = $"{result},{Adress}:{Formula}:{Value}:{rowId}";
                    rowIds = $"{rowIds};{cell.Start.Row}";

                }

                //Assert.AreEqual(";A1;A2;A3;A4;A5;A6;A7;A8;A9;A10;A11;A12", colRange);
                Assert.AreEqual(";A2;A3;A4;A5;A6;A7;A8;A10;A11;A12", colRange);
                //Assert.AreEqual(";1;2;3;4;5;6;7;8;9;10;11;12", rowIds);
                Assert.AreEqual(";2;3;4;5;6;7;8;10;11;12", rowIds);

                //Assert.AreEqual(12, counterColRange);
                Assert.AreEqual(10, counterColRange);

            }
        }

        [TestMethod]
        public void SumsIssue()
        {
            using (var p = new ExcelPackage())
            {
                p.Workbook.Worksheets.Add("first");

                var sheet = p.Workbook.Worksheets.First();

                var cellA1 = sheet.Cells["A1"];
                var cellA2 = sheet.Cells["A2"];
                var cellA3 = sheet.Cells["A3"];
                var cellA4 = sheet.Cells["A4"];
                var cellA5 = sheet.Cells["A5"];
                var cellA6 = sheet.Cells["A6"];
                var cellA7 = sheet.Cells["A7"];
                var cellA8 = sheet.Cells["A8"];
                var cellA9 = sheet.Cells["A9"];
                var cellA10 = sheet.Cells["A10"];
                var cellA11 = sheet.Cells["A11"];
                var cellA12 = sheet.Cells["A12"];
                cellA1.Value = 1;
                cellA2.Value = 1;
                cellA3.Value = 1;
                cellA4.Value = 1;
                cellA5.Value = 1;
                cellA6.Value = 1;
                cellA7.Value = 1;
                cellA8.Value = 1;
                cellA9.Value = 1;
                cellA10.Value = 1;
                cellA11.Value = 1;
                cellA12.Formula = "SUM(A1:A3,A5,A6,A7,A8,A10,A9,A11)";

                int counterFirstIteration = 0;
                int counterSecondIteration = 0;


                int CounterSingleAdress = 0;
                int CounterMultipleRanges = 0;
                int CounterRangesFirst = 0;
                int CounterRangesLast = 0;
                int counterNoRanges = 0;
                int counterOneRange = 0;
                int counterMixed = 0;
                var cellsFirstIteration = string.Empty;
                var cellsSecondIteration = string.Empty;
                var cellsSingleAdress = string.Empty;
                var cellsMultipleRanges = string.Empty;
                var cellsRangesFirst = string.Empty;
                var cellsRangesLast = string.Empty;
                var cellsNoRanges = string.Empty;
                var cellsOneRange = string.Empty;
                var cellsMixed = string.Empty;


                var range = sheet.Cells["A1:A3,A5,A6,A7,A8,A10,A9,A11"];
                foreach (var cell in range)
                {
                    counterFirstIteration++;
                    cellsFirstIteration = $"{cellsFirstIteration};{cell.Address}";
                }

                foreach (var cell in range)
                {
                    counterSecondIteration++;
                    cellsSecondIteration = $"{cellsSecondIteration};{cell.Address}";
                }

                Assert.AreEqual(cellsFirstIteration, cellsSecondIteration);
                Assert.AreEqual(cellsFirstIteration, ";A1;A2;A3;A5;A6;A7;A8;A10;A9;A11");

                Assert.AreEqual(counterFirstIteration, counterSecondIteration);
                Assert.AreEqual(10, counterFirstIteration);




                var rangeSingleAdress = sheet.Cells["A1"];
                foreach (var cell in rangeSingleAdress)
                {
                    CounterSingleAdress++;
                    cellsSingleAdress = $"{cellsSingleAdress};{cell.Address}";
                }

                Assert.AreEqual(";A1", cellsSingleAdress);
                Assert.AreEqual(1, CounterSingleAdress);

                cellsSingleAdress = String.Empty;
                CounterSingleAdress = 0;
                foreach (var cell in rangeSingleAdress)
                {
                    CounterSingleAdress++;
                    cellsSingleAdress = $"{cellsSingleAdress};{cell.Address}";
                }

                Assert.AreEqual(";A1", cellsSingleAdress);
                Assert.AreEqual(1, CounterSingleAdress);



                var rangeMultipleRanges = sheet.Cells["A1:A4,A5:A7,A8:A11"];
                foreach (var cell in rangeMultipleRanges)
                {
                    CounterMultipleRanges++;
                    cellsMultipleRanges = $"{cellsMultipleRanges};{cell.Address}";
                }

                Assert.AreEqual(";A1;A2;A3;A4;A5;A6;A7;A8;A9;A10;A11", cellsMultipleRanges);
                Assert.AreEqual(11, CounterMultipleRanges);

                CounterMultipleRanges = 0;
                cellsMultipleRanges = String.Empty;
                foreach (var cell in rangeMultipleRanges)
                {
                    CounterMultipleRanges++;
                    cellsMultipleRanges = $"{cellsMultipleRanges};{cell.Address}";
                }

                Assert.AreEqual(";A1;A2;A3;A4;A5;A6;A7;A8;A9;A10;A11", cellsMultipleRanges);
                Assert.AreEqual(11, CounterMultipleRanges);



                var rangeRangeFirst = sheet.Cells["A1:A4,A5,A6,A7"];
                foreach (var cell in rangeRangeFirst)
                {
                    CounterRangesFirst++;
                    cellsRangesFirst = $"{cellsRangesFirst};{cell.Address}";
                }

                Assert.AreEqual(";A1;A2;A3;A4;A5;A6;A7", cellsRangesFirst);
                Assert.AreEqual(7, CounterRangesFirst);

                CounterRangesFirst = 0;
                cellsRangesFirst = String.Empty;
                foreach (var cell in rangeRangeFirst)
                {
                    CounterRangesFirst++;
                    cellsRangesFirst = $"{cellsRangesFirst};{cell.Address}";
                }

                Assert.AreEqual(";A1;A2;A3;A4;A5;A6;A7", cellsRangesFirst);
                Assert.AreEqual(7, CounterRangesFirst);



                var rangeRangeLast = sheet.Cells["A1,A2,A3,A4:A7"];
                foreach (var cell in rangeRangeLast)
                {
                    CounterRangesLast++;
                    cellsRangesLast = $"{cellsRangesLast};{cell.Address}";
                }

                Assert.AreEqual(";A1;A2;A3;A4;A5;A6;A7", cellsRangesLast);
                Assert.AreEqual(7, CounterRangesLast);

                CounterRangesLast = 0;
                cellsRangesLast = String.Empty;
                foreach (var cell in rangeRangeLast)
                {
                    CounterRangesLast++;
                    cellsRangesLast = $"{cellsRangesLast};{cell.Address}";
                }

                Assert.AreEqual(";A1;A2;A3;A4;A5;A6;A7", cellsRangesLast);
                Assert.AreEqual(7, CounterRangesLast);



                var rangeOneRange = sheet.Cells["A1:A7"];
                foreach (var cell in rangeOneRange)
                {
                    counterOneRange++;
                    cellsOneRange = $"{cellsOneRange};{cell.Address}";
                }

                Assert.AreEqual(";A1;A2;A3;A4;A5;A6;A7", cellsOneRange);
                Assert.AreEqual(7, counterOneRange);

                counterOneRange = 0;
                cellsOneRange = String.Empty;
                foreach (var cell in rangeOneRange)
                {
                    counterOneRange++;
                    cellsOneRange = $"{cellsOneRange};{cell.Address}";
                }

                Assert.AreEqual(";A1;A2;A3;A4;A5;A6;A7", cellsOneRange);
                Assert.AreEqual(7, counterOneRange);



                var rangeNoRange = sheet.Cells["A1,A2,A3,A4"];
                foreach (var cell in rangeNoRange)
                {
                    counterNoRanges++;
                    cellsNoRanges = $"{cellsNoRanges};{cell.Address}";
                }

                Assert.AreEqual(";A1;A2;A3;A4", cellsNoRanges);
                Assert.AreEqual(4, counterNoRanges);

                counterNoRanges = 0;
                cellsNoRanges = String.Empty;
                foreach (var cell in rangeNoRange)
                {
                    counterNoRanges++;
                    cellsNoRanges = $"{cellsNoRanges};{cell.Address}";
                }

                Assert.AreEqual(";A1;A2;A3;A4", cellsNoRanges);
                Assert.AreEqual(4, counterNoRanges);



                var rangeMixed = sheet.Cells["A1,A2,A3:A5,A6,A7"];
                foreach (var cell in rangeMixed)
                {
                    counterMixed++;
                    cellsMixed = $"{cellsMixed};{cell.Address}";
                }

                Assert.AreEqual(";A1;A2;A3;A4;A5;A6;A7", cellsMixed);
                Assert.AreEqual(7, counterMixed);

                counterMixed = 0;
                cellsMixed = String.Empty;
                foreach (var cell in rangeMixed)
                {
                    counterMixed++;
                    cellsMixed = $"{cellsMixed};{cell.Address}";
                }

                Assert.AreEqual(";A1;A2;A3;A4;A5;A6;A7", cellsMixed);
                Assert.AreEqual(7, counterMixed);



                int counter = 0;
                String cells = String.Empty;
                foreach (var cell in sheet.Cells)
                {
                    counter++;
                    cells = $"{cells};{cell.Address}";
                }

                Assert.AreEqual(";A1;A2;A3;A4;A5;A6;A7;A8;A9;A10;A11;A12", cells);
                Assert.AreEqual(12, counter);

                counter = 0;
                cells = String.Empty;
                foreach (var cell in sheet.Cells)
                {
                    counter++;
                    cells = $"{cells};{cell.Address}";
                }

                Assert.AreEqual(";A1;A2;A3;A4;A5;A6;A7;A8;A9;A10;A11;A12", cells);
                Assert.AreEqual(12, counter);



                int counterColRange = 0;
                var colRange = String.Empty;
                foreach (var cell in sheet.Cells["A:A"])
                {
                    counterColRange++;
                    colRange = $"{colRange};{cell.Address}";
                }

                Assert.AreEqual(";A1;A2;A3;A4;A5;A6;A7;A8;A9;A10;A11;A12", colRange);
                Assert.AreEqual(12, counterColRange);

                counterColRange = 0;
                colRange = String.Empty;
                foreach (var cell in sheet.Cells["A:A"])
                {
                    counterColRange++;
                    colRange = $"{colRange};{cell.Address}";
                }

                Assert.AreEqual(";A1;A2;A3;A4;A5;A6;A7;A8;A9;A10;A11;A12", colRange);
                Assert.AreEqual(12, counterColRange);





                // var cellA13 = sheet.Cells["A13"];
                var cellB13 = sheet.Cells["B13"];
                var cellC13 = sheet.Cells["C13"];
                var cellD13 = sheet.Cells["D13"];
                var cellE13 = sheet.Cells["E13"];
                var cellF13 = sheet.Cells["F13"];
                //var cellG13 = sheet.Cells["G13"];
                var cellH13 = sheet.Cells["H13"];
                var cellI13 = sheet.Cells["I13"];
                var cellJ13 = sheet.Cells["J13"];
                var cellK13 = sheet.Cells["K13"];
                var cellL13 = sheet.Cells["L13"];
                //cellA13.Value = 1;
                cellB13.Value = 1;
                cellC13.Value = 1;
                cellD13.Value = 1;
                cellE13.Value = 1;
                cellF13.Value = 1;
                //cellG13.Value = 1;
                cellH13.Value = 1;
                cellI13.Value = 1;
                cellJ13.Value = 1;
                cellK13.Value = 1;
                cellL13.Formula = "SUM(A1:A3,A5,A6,A7,A8,A10,A9,A11)";

                int counterHorizontal = 0;
                String rows = String.Empty;
                foreach (var cell in sheet.Cells["13:13"])
                {
                    counterHorizontal++;
                    rows = $"{rows};{cell.Address}";
                }

                Assert.AreEqual(";B13;C13;D13;E13;F13;H13;I13;J13;K13;L13", rows);
                Assert.AreEqual(10, counterHorizontal);

                counterHorizontal = 0;
                rows = String.Empty;
                foreach (var cell in sheet.Cells["13:13"])
                {
                    counterHorizontal++;
                    rows = $"{rows};{cell.Address}";
                }

                Assert.AreEqual(";B13;C13;D13;E13;F13;H13;I13;J13;K13;L13", rows);
                Assert.AreEqual(10, counterHorizontal);

            }
        }

        [TestMethod]
        public void Issue_WithNamedRanges()
        {
            var excelTestFile = Resources.TestDoc_NamedRangeInFormula_xlsx;
            using (MemoryStream excelStream = new MemoryStream())
            {
                excelStream.Write(excelTestFile, 0, excelTestFile.Length);

                using (ExcelPackage exlPackage = new ExcelPackage(excelStream))
                {
                    exlPackage.Workbook.Worksheets.First().Cells["A1"].Value = "Changed Value";
                    exlPackage.Workbook.Calculate();

                    Assert.AreEqual("Changed Value",
                        exlPackage.Workbook.Worksheets.First().Cells["A4"].Value.ToString());
                }
            }
        }

        [TestMethod]
        public void Issue_WithRangeCalculation()
        {
            //Issue: If two namedRanges (columns with Names) are calculated like "=range1 + range2" Only the first row of the ranges are calculated and the result is copied to the rest of the rows from the resultcolumn. 
            var excelTestFile = Resources.Issue_WithRangeCalculation;
            using (MemoryStream excelStream = new MemoryStream())
            {
                excelStream.Write(excelTestFile, 0, excelTestFile.Length);

                using (ExcelPackage exlPackage = new ExcelPackage(excelStream))
                {
                    var ws = exlPackage.Workbook.Worksheets[1];
                    ws.Calculate();

                    //range in range in Fomular
                    Assert.AreEqual(311d, ws.Cells["C1"].Value);
                    Assert.AreEqual(306d, ws.Cells["C2"].Value);

                    //range1+range2 horizontal
                    Assert.AreEqual(103d, ws.Cells["C3"].Value);
                    Assert.AreEqual(104d, ws.Cells["C4"].Value);
                    Assert.AreEqual(105d, ws.Cells["C5"].Value);
                    Assert.AreEqual(106d, ws.Cells["C6"].Value);
                    Assert.AreEqual(107d, ws.Cells["C7"].Value);
                    Assert.AreEqual(108d, ws.Cells["C8"].Value);
                    Assert.AreEqual(109d, ws.Cells["C9"].Value);
                    Assert.AreEqual(110d, ws.Cells["C10"].Value);

                    Assert.AreEqual(112d, ws.Cells["C12"].Value);
                    Assert.AreEqual(113d, ws.Cells["C13"].Value);
                    Assert.AreEqual(114d, ws.Cells["C14"].Value);

                    //range3+range4 vertical
                    Assert.AreEqual(101d, ws.Cells["F21"].Value);
                    Assert.AreEqual(102d, ws.Cells["G21"].Value);
                    Assert.AreEqual(103d, ws.Cells["H21"].Value);
                    Assert.AreEqual(104d, ws.Cells["I21"].Value);
                    Assert.AreEqual(105d, ws.Cells["J21"].Value);
                    Assert.AreEqual(106d, ws.Cells["K21"].Value);
                    Assert.AreEqual(107d, ws.Cells["L21"].Value);
                    Assert.AreEqual(108d, ws.Cells["M21"].Value);
                    Assert.AreEqual(109d, ws.Cells["N21"].Value);
                    Assert.AreEqual(110d, ws.Cells["O21"].Value);
                    Assert.AreEqual(111d, ws.Cells["P21"].Value);
                    Assert.AreEqual(112d, ws.Cells["Q21"].Value);
                    Assert.AreEqual(113d, ws.Cells["R21"].Value);

                    //When Issue_WithRangeCalculation_IF
                    Assert.AreEqual(306d, ws.Cells["H2"].Value);
                    Assert.AreEqual(103d, ws.Cells["H3"].Value);
                    Assert.AreEqual(104d, ws.Cells["H4"].Value);
                    Assert.AreEqual(105d, ws.Cells["H5"].Value);

                    Assert.AreEqual(100d, ws.Cells["I2"].Value);
                    Assert.AreEqual(100d, ws.Cells["I3"].Value);
                    Assert.AreEqual(100d, ws.Cells["I4"].Value);
                    Assert.AreEqual(100d, ws.Cells["I5"].Value);

                    Assert.AreEqual(100d, ws.Cells["J2"].Value);
                    Assert.AreEqual(100d, ws.Cells["J3"].Value);
                    Assert.AreEqual(100d, ws.Cells["J4"].Value);
                    Assert.AreEqual(100d, ws.Cells["J5"].Value);

                    Assert.AreEqual("Falsche Auswahl", ws.Cells["K2"].Value);
                    Assert.AreEqual("Falsche Auswahl", ws.Cells["K3"].Value);
                    Assert.AreEqual("Falsche Auswahl", ws.Cells["K4"].Value);
                    Assert.AreEqual("Falsche Auswahl", ws.Cells["K5"].Value);


                    //Normal
                    Assert.AreEqual(198d, ws.Cells["C18"].Value);

                    //String
                    Assert.AreEqual("#VALUE!", ws.Cells["C19"].Value.ToString());
                    Assert.AreEqual("#VALUE!", ws.Cells["C15"].Value.ToString());

                    //Empty Cell
                    Assert.AreEqual(100d, ws.Cells["C11"].Value);
                    Assert.AreEqual(20d, ws.Cells["C20"].Value);

                    //OutOfRange IF
                    Assert.AreEqual("#VALUE!", ws.Cells["H1"].Value.ToString());
                    Assert.AreEqual("#VALUE!", ws.Cells["I1"].Value.ToString());
                    Assert.AreEqual("#VALUE!", ws.Cells["J1"].Value.ToString());
                    Assert.AreEqual("Falsche Auswahl", ws.Cells["K1"].Value);
                    Assert.AreEqual("#VALUE!", ws.Cells["H6"].Value.ToString());
                    Assert.AreEqual("#VALUE!", ws.Cells["I6"].Value.ToString());
                    Assert.AreEqual("#VALUE!", ws.Cells["J6"].Value.ToString());
                    Assert.AreEqual("Falsche Auswahl", ws.Cells["K6"].Value);

                    //OutOfRange Normal
                    Assert.AreEqual("#VALUE!", ws.Cells["C16"].Value.ToString());
                    Assert.AreEqual("#VALUE!", ws.Cells["E21"].Value.ToString());
                    Assert.AreEqual("#VALUE!", ws.Cells["S21"].Value.ToString());

                    //UseAGAIN
                    Assert.AreEqual(206d, ws.Cells["F2"].Value);
                    Assert.AreEqual(3d, ws.Cells["F3"].Value);
                    Assert.AreEqual(4d, ws.Cells["F4"].Value);
                    Assert.AreEqual(5d, ws.Cells["F5"].Value);
                    //UseIFAGAIN
                    Assert.AreEqual(306d, ws.Cells["M2"].Value);
                    Assert.AreEqual(103d, ws.Cells["M3"].Value);
                    Assert.AreEqual(104d, ws.Cells["M4"].Value);
                    Assert.AreEqual(105d, ws.Cells["M5"].Value);
                    Assert.AreEqual("#VALUE!", ws.Cells["M6"].Value.ToString());


                    //Check if something in if is fixed wrong
                    Assert.AreEqual(2d, ws.Cells["F11"].Value);
                    Assert.AreEqual(1d, ws.Cells["F12"].Value);

                }
            }
        }


        [TestMethod]
        public void IssueWithExternalFormulas()
        {
            //Issue: If a formula contains external links the old value should be used instead of resulting in #NAME-Error
            var excelTestFile = Resources.ExternalReferences;
            using (MemoryStream excelStream = new MemoryStream())
            {
                excelStream.Write(excelTestFile, 0, excelTestFile.Length);

                using (ExcelPackage exlPackage = new ExcelPackage(excelStream))
                {
                    var ws = exlPackage.Workbook.Worksheets[1];
                    ws.Calculate();

                    Assert.AreEqual(60d, ws.Cells["A1"].Value);
                    Assert.AreEqual(60d, ws.Cells["A2"].Value);
                    Assert.AreEqual(23d, ws.Cells["B19"].Value);
                    Assert.AreEqual(23d, ws.Cells["B20"].Value);



                }
            }
        }


        [TestMethod]
        public void IssueWithChangedCostumstyles()
        {
            //Issue: If the Custom / Built-in Styles are Saved not in EPPlus the information about their property gets lost 
            var excelTestFile = Resources.Layout_Format_vorlage;
            using (MemoryStream excelStream = new MemoryStream())
            {
                excelStream.Write(excelTestFile, 0, excelTestFile.Length);

                using (ExcelPackage exlPackage = new ExcelPackage(excelStream))
                {

                    var styles = exlPackage.Workbook.Styles;
                    Assert.AreEqual(0, styles.CellStyleXfs[0].NumberFormatId);
                    Assert.AreEqual(0, styles.CellStyleXfs[0].FontId);
                    Assert.AreEqual(0, styles.CellStyleXfs[0].FillId);
                    Assert.AreEqual(0, styles.CellStyleXfs[0].BorderId);
                    Assert.IsNull(styles.CellStyleXfs[0].ApplyNumberFormat);
                    Assert.IsNull(styles.CellStyleXfs[0].ApplyFill);
                    Assert.IsNull(styles.CellStyleXfs[0].ApplyBorder);
                    Assert.IsNull(styles.CellStyleXfs[0].ApplyAlignment);
                    Assert.IsNull(styles.CellStyleXfs[0].ApplyProtection);

                    Assert.AreEqual(0, styles.CellStyleXfs[1].NumberFormatId);
                    Assert.AreEqual(1, styles.CellStyleXfs[1].FontId);
                    Assert.AreEqual(0, styles.CellStyleXfs[1].FillId);
                    Assert.AreEqual(0, styles.CellStyleXfs[1].BorderId);
                    Assert.AreEqual(false, styles.CellStyleXfs[1].ApplyNumberFormat);
                    Assert.AreEqual(false, styles.CellStyleXfs[1].ApplyFill);
                    Assert.AreEqual(false, styles.CellStyleXfs[1].ApplyBorder);
                    Assert.AreEqual(false, styles.CellStyleXfs[1].ApplyAlignment);
                    Assert.AreEqual(false, styles.CellStyleXfs[1].ApplyProtection);

                    Assert.AreEqual(0, styles.CellXfs[0].NumberFormatId);
                    Assert.AreEqual(0, styles.CellXfs[0].FontId);
                    Assert.AreEqual(0, styles.CellXfs[0].FillId);
                    Assert.AreEqual(0, styles.CellXfs[0].BorderId);
                    Assert.IsNull(styles.CellXfs[0].ApplyNumberFormat);
                    Assert.IsNull(styles.CellXfs[0].ApplyFill);
                    Assert.IsNull(styles.CellXfs[0].ApplyBorder);
                    Assert.IsNull(styles.CellXfs[0].ApplyAlignment);
                    Assert.IsNull(styles.CellXfs[0].ApplyProtection);

                    Assert.AreEqual(0, styles.CellXfs[1].NumberFormatId);
                    Assert.AreEqual(1, styles.CellXfs[1].FontId);
                    Assert.AreEqual(0, styles.CellXfs[1].FillId);
                    Assert.AreEqual(0, styles.CellXfs[1].BorderId);
                    Assert.IsNull(styles.CellXfs[1].ApplyNumberFormat);
                    Assert.IsNull(styles.CellXfs[1].ApplyFill);
                    Assert.IsNull(styles.CellXfs[1].ApplyBorder);
                    Assert.IsNull(styles.CellXfs[1].ApplyAlignment);
                    Assert.IsNull(styles.CellXfs[1].ApplyProtection);

                }

            }

        }

        [TestMethod]
        public void IssueWithRoman()
        {
            //Issue: Roman is not implementet
            var excelTestFile = Resources.Roman_allNumbers;
            using (MemoryStream excelStream = new MemoryStream())
            {
                excelStream.Write(excelTestFile, 0, excelTestFile.Length);

                using (ExcelPackage exlPackage = new ExcelPackage(excelStream))
                {
                    var ws = exlPackage.Workbook.Worksheets[1];

                    ws.Calculate();

                    //no Parameter
                    for (int i = 1; i <= ws.Cells["A:A"].Count(); i++)
                        Assert.AreEqual(ws.Cells[i, 1].Value, ws.Cells[i, (1 + 11)].Value);

                    //Parameter 0
                    for (int i = 1; i <= ws.Cells["B:B"].Count(); i++)
                        Assert.AreEqual(ws.Cells[i, 2].Value, ws.Cells[i, (2 + 11)].Value);
                    //Parameter 1
                    for (int i = 1; i <= ws.Cells["C:C"].Count(); i++)
                        Assert.AreEqual(ws.Cells[i, 3].Value, ws.Cells[i, (3 + 11)].Value);
                    //Parameter 2
                    for (int i = 1; i <= ws.Cells["D:D"].Count(); i++)
                        Assert.AreEqual(ws.Cells[i, 4].Value, ws.Cells[i, (4 + 11)].Value);
                    //Parameter 3
                    for (int i = 1; i <= ws.Cells["E:E"].Count(); i++)
                        Assert.AreEqual(ws.Cells[i, 5].Value, ws.Cells[i, (5 + 11)].Value);
                    //Parameter 4
                    for (int i = 1; i <= ws.Cells["F:F"].Count(); i++)
                        Assert.AreEqual(ws.Cells[i, 6].Value, ws.Cells[i, (6 + 11)].Value);
                    //Parameter TRUE
                    for (int i = 1; i <= ws.Cells["G:G"].Count(); i++)
                        Assert.AreEqual(ws.Cells[i, 7].Value, ws.Cells[i, (7 + 11)].Value);
                    //Parameter FALSE
                    for (int i = 1; i <= ws.Cells["H:H"].Count(); i++)
                        Assert.AreEqual(ws.Cells[i, 7].Value, ws.Cells[i, (7 + 11)].Value);

                }
            }
        }

        [TestMethod]
        public void IssueWithRomanSMALL()
        {
            //Issue: Roman is not implementet
            var excelTestFile = Resources.Roman;
            using (MemoryStream excelStream = new MemoryStream())
            {
                excelStream.Write(excelTestFile, 0, excelTestFile.Length);

                using (ExcelPackage exlPackage = new ExcelPackage(excelStream))
                {
                    var ws = exlPackage.Workbook.Worksheets[1];

                    ws.Calculate();

                    //Parameter
                    Assert.AreEqual(ws.Cells["A1"].Value, ws.Cells["B1"].Value);
                    Assert.AreEqual(ws.Cells["A2"].Value, ws.Cells["B2"].Value);
                    Assert.AreEqual(ws.Cells["A3"].Value, ws.Cells["B3"].Value);
                    Assert.AreEqual(ws.Cells["A4"].Value, ws.Cells["B4"].Value);
                    Assert.AreEqual(ws.Cells["A5"].Value, ws.Cells["B5"].Value);
                    Assert.AreEqual(ws.Cells["A6"].Value, ws.Cells["B6"].Value);
                    Assert.AreEqual(ws.Cells["A7"].Value, ws.Cells["B7"].Value);
                    Assert.AreEqual(ws.Cells["A8"].Value, ws.Cells["B8"].Value);
                    Assert.AreEqual(ws.Cells["A9"].Value, ws.Cells["B9"].Value);
                    Assert.AreEqual(ws.Cells["A10"].Value, ws.Cells["B10"].Value);


                    //wrong Parameter
                    Assert.AreEqual(ws.Cells["C1"].Value, ws.Cells["D1"].Value);
                    Assert.AreEqual(ws.Cells["C2"].Value, ws.Cells["D2"].Value);
                    Assert.AreEqual(ws.Cells["C3"].Value, ws.Cells["D3"].Value);
                    Assert.AreEqual(ws.Cells["C4"].Value, ws.Cells["D4"].Value);
                    Assert.AreEqual(ws.Cells["C5"].Value, String.Empty);
                }
            }
        }

        [TestMethod]
        public void IssueWithTrim()
        {
            //Issue: Trim is not implementet
            var excelTestFile = Resources.Trim;
            using (MemoryStream excelStream = new MemoryStream())
            {
                excelStream.Write(excelTestFile, 0, excelTestFile.Length);

                using (ExcelPackage exlPackage = new ExcelPackage(excelStream))
                {
                    var ws = exlPackage.Workbook.Worksheets[2];

                    ws.Calculate();

                    //Parameter
                    Assert.AreEqual("Anlagevermögen", ws.Cells["B8"].Value);
                    Assert.AreEqual("123 456 ABC", ws.Cells["B9"].Value);

                }
            }
        }


        [TestMethod]
        public void Issue15347()
        {
            using (var file = new MemoryStream())
            {
                file.Write(EPPlusTest.Properties.Resources.Issue15347Test, 0,
                    EPPlusTest.Properties.Resources.Issue15347Test.Length);
                using (ExcelPackage package = new ExcelPackage(file))
                {

                    var sheet = package.Workbook.Worksheets[1];

                    ExcelRange range = sheet.Cells[1, 1];
                    package.Save();
                }

                using (ExcelPackage package = new ExcelPackage(file))
                {
                    var sheet = package.Workbook.Worksheets["Tabelle1"];
                    ExcelRange range = sheet.Cells[1, 1];
                    Assert.IsTrue(range.Value is string);
                }
            }
        }


        [TestMethod]
        public void Issue15353_QuotsInNamedRanges()
        {
            var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("My 'Sheet");
            var namedRange1 = package.Workbook.Names.Add("name1", new ExcelRangeBase(sheet, "$B$1"));

            Assert.AreEqual(namedRange1.FullAddressAbsolute, "'My ''Sheet'!$B$1");


            //"My "Sheet"            //'My "Sheet'!$A$1

        }

        [TestMethod]
        public void Issue15353_QuotsInFormulae()
        {
            //positive Example 1:
            TokenizerContext formulaContext1 = new TokenizerContext("= part1 & \" is \'\" & yearshort");

            formulaContext1.ToggleIsInSheetName();

            TokenHandler handlerSheetName1 =
                new TokenHandler(formulaContext1, _tokenFactory, new TokenSeparatorProvider());
            handlerSheetName1.Worksheet = "My 'Sheet";

            while (handlerSheetName1.HasMore())
            {
                handlerSheetName1.Next();
            }

            var _syntAnalyzer = new SyntacticAnalyzer();

            try
            {
                _syntAnalyzer.Analyze(formulaContext1.Result);

            }
            catch
            {
                Assert.IsTrue(false, "Syntactic analyzer throws exception on a valid string");
            }


            //positive Example 2:

            TokenizerContext formulaContext2 = new TokenizerContext("= part1 & \" is \"\"some\"\" \" & yearshort");

            formulaContext2.ToggleIsInSheetName();

            TokenHandler handlerSheetName2 =
                new TokenHandler(formulaContext2, _tokenFactory, new TokenSeparatorProvider());
            handlerSheetName2.Worksheet = "My 'Sheet";

            while (handlerSheetName2.HasMore())
            {
                handlerSheetName2.Next();
            }

            var _syntAnalyzer2 = new SyntacticAnalyzer();

            try
            {
                _syntAnalyzer2.Analyze(formulaContext2.Result);

            }
            catch
            {
                Assert.IsTrue(false, "Syntactic analyzer throws exception on a valid string");
            }


            //positive Example 3:

            TokenizerContext formulaContext3 = new TokenizerContext("= part1 & \" is some \"\"\" & yearshort");

            formulaContext3.ToggleIsInSheetName();

            TokenHandler handlerSheetName3 =
                new TokenHandler(formulaContext3, _tokenFactory, new TokenSeparatorProvider());
            handlerSheetName3.Worksheet = "My 'Sheet";

            while (handlerSheetName3.HasMore())
            {
                handlerSheetName3.Next();
            }

            var _syntAnalyzer3 = new SyntacticAnalyzer();

            try
            {
                _syntAnalyzer3.Analyze(formulaContext3.Result);
            }
            catch
            {
                Assert.IsTrue(false, "Syntactic analyzer throws exception on a valid string");
            }


            //positive Example 4:

            TokenizerContext formulaContext4 = new TokenizerContext("= part1 & \"\' is some\'\" & yearshort");

            formulaContext4.ToggleIsInSheetName();

            TokenHandler handlerSheetName4 =
                new TokenHandler(formulaContext4, _tokenFactory, new TokenSeparatorProvider());
            handlerSheetName4.Worksheet = "My 'Sheet";

            while (handlerSheetName4.HasMore())
            {
                handlerSheetName4.Next();
            }

            var _syntAnalyzer4 = new SyntacticAnalyzer();

            try
            {
                _syntAnalyzer4.Analyze(formulaContext4.Result);

            }
            catch
            {
                Assert.IsTrue(false, "Syntactic analyzer throws exception on a valid string");
            }


            //positive Example 5:

            TokenizerContext formulaContext5 = new TokenizerContext("= part1 & \" is some\'\'\" & yearshort");

            formulaContext5.ToggleIsInSheetName();

            TokenHandler handlerSheetName5 =
                new TokenHandler(formulaContext5, _tokenFactory, new TokenSeparatorProvider());
            handlerSheetName5.Worksheet = "My 'Sheet";

            while (handlerSheetName5.HasMore())
            {
                handlerSheetName5.Next();
            }

            var _syntAnalyzer5 = new SyntacticAnalyzer();

            try
            {
                _syntAnalyzer5.Analyze(formulaContext5.Result);
            }
            catch
            {
                Assert.IsTrue(false, "Syntactic analyzer throws exception on a valid string");
            }


            //negative Example 1:

            TokenizerContext formulaContext6 = new TokenizerContext("= part1 & \" is \"\"some\" \" & yearshort");


            formulaContext6.ToggleIsInSheetName();

            TokenHandler handlerSheetName6 =
                new TokenHandler(formulaContext6, _tokenFactory, new TokenSeparatorProvider());
            handlerSheetName6.Worksheet = "My 'Sheet";

            while (handlerSheetName6.HasMore())
            {
                handlerSheetName6.Next();
            }

            var _syntAnalyzer6 = new SyntacticAnalyzer();

            try
            {
                _syntAnalyzer6.Analyze(formulaContext6.Result);
                Assert.Fail("Syntactic analyzer has to throw an exception on an invalid string");
            }
            catch
            {
            }
            //negative Example 2: 

            TokenizerContext formulaContext7 = new TokenizerContext("= part1 & \" is some \"\" & yearshort");

            formulaContext7.ToggleIsInSheetName();

            TokenHandler handlerSheetName7 =
                new TokenHandler(formulaContext7, _tokenFactory, new TokenSeparatorProvider());
            handlerSheetName7.Worksheet = "My 'Sheet";

            while (handlerSheetName7.HasMore())
            {
                handlerSheetName7.Next();
            }

            var _syntAnalyzer7 = new SyntacticAnalyzer();

            try
            {
                _syntAnalyzer7.Analyze(formulaContext7.Result);
                Assert.Fail("Syntactic analyzer has to throw an exception on an invalid string");
            }
            catch
            {
            }

            //negative Example 3:

            TokenizerContext formulaContext8 = new TokenizerContext("= part1 & \"\' is some\' & yearshort");

            //negative
            //= part1 & " is some'' & yearshort


            formulaContext8.ToggleIsInSheetName();

            TokenHandler handlerSheetName8 =
                new TokenHandler(formulaContext8, _tokenFactory, new TokenSeparatorProvider());
            handlerSheetName8.Worksheet = "My 'Sheet";

            while (handlerSheetName8.HasMore())
            {
                handlerSheetName8.Next();
            }

            var _syntAnalyzer8 = new SyntacticAnalyzer();

            try
            {
                _syntAnalyzer8.Analyze(formulaContext8.Result);
                Assert.Fail("Syntactic analyzer has to throw an exception on an invalid string");
            }
            catch
            {

            }


            //negative Example 4: 

            TokenizerContext formulaContext9 = new TokenizerContext("= part1 & \" is some\'\' & yearshort");

            formulaContext9.ToggleIsInSheetName();

            TokenHandler handlerSheetName9 =
                new TokenHandler(formulaContext9, _tokenFactory, new TokenSeparatorProvider());
            handlerSheetName9.Worksheet = "My 'Sheet";

            while (handlerSheetName9.HasMore())
            {
                handlerSheetName9.Next();
            }

            var _syntAnalyzer9 = new SyntacticAnalyzer();

            try
            {
                _syntAnalyzer9.Analyze(formulaContext9.Result);
                Assert.Fail("Syntactic analyzer has to throw an exception on an invalid string");
            }
            catch
            {

            }

        }

        [TestMethod]
        public void canNotSaveWithFileInfoAfterSavingPackageAsStream()
        {
            //If we save the file As a Stream and then saving it again an Index Out Of Range Exception is thrown. (Nachtrag Ticket #62877)
            ExcelPackage packageNew = new ExcelPackage();
            packageNew.Workbook.Worksheets.Add("sheet_1");
            

            using (var stream = new MemoryStream())
            {
                packageNew.SaveAs(stream);
            }

            packageNew.Save();


        }

        [TestMethod]
        public void IssueGermanBuildInNumberFormat()
        {
            //Issue: The German BuildInNumberFormat differs from the English BuildInNumberformat therefore Epplus has to check the culture before parsing the id to NumberFormatExpression.
            var excelTestFile = Resources.GermanBuildInNumberFormat;
            using (MemoryStream excelStream = new MemoryStream())
            {
                excelStream.Write(excelTestFile, 0, excelTestFile.Length);

                using (ExcelPackage exlPackage = new ExcelPackage(excelStream))
                {
                    var ws = exlPackage.Workbook.Worksheets[1];

                    if (System.Threading.Thread.CurrentThread.CurrentCulture.Name.Equals("de-DE"))
                    {

                        var excelFormatString_2 = ws.Cells[2, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("General", excelFormatString_2);

                        var excelFormatString_3 = ws.Cells[3, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("0", excelFormatString_3);

                        var excelFormatString_4 = ws.Cells[4, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("0.00", excelFormatString_4);

                        var excelFormatString_5 = ws.Cells[5, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("#,##0", excelFormatString_5);

                        var excelFormatString_6 = ws.Cells[6, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("#,##0.00", excelFormatString_6);

                        var excelFormatString_7 = ws.Cells[7, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("#,##0 _€;-#,##0 _€", excelFormatString_7);

                        var excelFormatString_8 = ws.Cells[8, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("#,##0 _€;[Red]-#,##0 _€", excelFormatString_8);

                        var excelFormatString_9 = ws.Cells[9, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("#,##0.00 _€;-#,##0.00 _€", excelFormatString_9);

                        var excelFormatString_10 = ws.Cells[10, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("#,##0.00 _€;[Red]-#,##0.00 _€", excelFormatString_10);

                        var excelFormatString_11 = ws.Cells[11, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("#,##0\\ \"€\";\\-#,##0\\ \"€\"", excelFormatString_11);

                        var excelFormatString_12 = ws.Cells[12, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("#,##0\\ \"€\";[Red]\\-#,##0\\ \"€\"", excelFormatString_12);

                        var excelFormatString_13 = ws.Cells[13, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("#,##0.00\\ \"€\";\\-#,##0.00\\ \"€\"", excelFormatString_13);

                        var excelFormatString_14 = ws.Cells[14, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("#,##0.00\\ \"€\";[Red]\\-#,##0.00\\ \"€\"", excelFormatString_14);

                        var excelFormatString_15 = ws.Cells[15, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("0%", excelFormatString_15);

                        var excelFormatString_16 = ws.Cells[16, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("0.00%", excelFormatString_16);

                        var excelFormatString_17 = ws.Cells[17, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("0.00E+00", excelFormatString_17);

                        var excelFormatString_18 = ws.Cells[18, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("##0.0E+0", excelFormatString_18);

                        var excelFormatString_19 = ws.Cells[19, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("# ?/?", excelFormatString_19);

                        var excelFormatString_20 = ws.Cells[20, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("# ??/??", excelFormatString_20);

                        var excelFormatString_21 = ws.Cells[21, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("dd.mm.yyyy", excelFormatString_21);

                        var excelFormatString_22 = ws.Cells[22, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("dd. mm yy", excelFormatString_22);

                        var excelFormatString_23 = ws.Cells[23, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("dd. mmm", excelFormatString_23);

                        var excelFormatString_24 = ws.Cells[24, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("mmm yy", excelFormatString_24);

                        var excelFormatString_25 = ws.Cells[25, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("h:mm AM/PM", excelFormatString_25);

                        var excelFormatString_26 = ws.Cells[26, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("h:mm:ss AM/PM", excelFormatString_26);

                        var excelFormatString_27 = ws.Cells[27, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("hh:mm", excelFormatString_27);

                        var excelFormatString_28 = ws.Cells[28, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("hh:mm:ss", excelFormatString_28);

                        var excelFormatString_29 = ws.Cells[29, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("dd.mm.yyyy hh:mm", excelFormatString_29);

                        var excelFormatString_30 = ws.Cells[30, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("mm:ss", excelFormatString_30);

                        var excelFormatString_31 = ws.Cells[31, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("mm:ss.0", excelFormatString_31);

                        var excelFormatString_32 = ws.Cells[32, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("@", excelFormatString_32);

                        var excelFormatString_33 = ws.Cells[33, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("[h]:mm:ss", excelFormatString_33);

                        var excelFormatString_34 = ws.Cells[34, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("_-* #,##0\\ \"€\"_-;\\-* #,##0\\ \"€\"_-;_-* \"-\"\\ \"€\"_-;_-@_-",
                            excelFormatString_34);

                        var excelFormatString_35 = ws.Cells[35, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("_-* #,##0\\ _€_-;\\-* #,##0\\ _€_-;_-* \"-\"\\ _€_-;_-@_-",
                            excelFormatString_35);

                        var excelFormatString_36 = ws.Cells[36, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("_-* #,##0.00\\ \"€\"_-;\\-* #,##0.00\\ \"€\"_-;_-* \"-\"??\\ \"€\"_-;_-@_-",
                            excelFormatString_36);

                        var excelFormatString_37 = ws.Cells[37, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("_-* #,##0.00\\ _€_-;\\-* #,##0.00\\ _€_-;_-* \"-\"??\\ _€_-;_-@_-",
                            excelFormatString_37);

                        var excelFormatString_38 = ws.Cells[38, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("mmm\\ yyyy", excelFormatString_38);

                        var excelFormatString_39 = ws.Cells[39, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("[$-407]dddd\\,\\ d/\\ mmmm\\ yyyy", excelFormatString_39);

                    }
                    else
                    {
                        //if you want to check the general way, please change your windowslanguagesettings to English.

                        var excelFormatString_2 = ws.Cells[2, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("General", excelFormatString_2);

                        var excelFormatString_3 = ws.Cells[3, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("0", excelFormatString_3);

                        var excelFormatString_4 = ws.Cells[4, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("0.00", excelFormatString_4);

                        var excelFormatString_5 = ws.Cells[5, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("#,##0", excelFormatString_5);

                        var excelFormatString_6 = ws.Cells[6, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("#,##0.00", excelFormatString_6);

                        var excelFormatString_7 = ws.Cells[7, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("#,##0 ;(#,##0)", excelFormatString_7);

                        var excelFormatString_8 = ws.Cells[8, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("#,##0 ;[Red](#,##0)", excelFormatString_8);

                        var excelFormatString_9 = ws.Cells[9, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("#,##0.00;(#,##0.00)", excelFormatString_9);

                        var excelFormatString_10 = ws.Cells[10, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("#,##0.00;[Red](#,##0.00)", excelFormatString_10);

                        var excelFormatString_11 = ws.Cells[11, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("#,##0\\ \"€\";\\-#,##0\\ \"€\"", excelFormatString_11);

                        var excelFormatString_12 = ws.Cells[12, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("#,##0\\ \"€\";[Red]\\-#,##0\\ \"€\"", excelFormatString_12);

                        var excelFormatString_13 = ws.Cells[13, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("#,##0.00\\ \"€\";\\-#,##0.00\\ \"€\"", excelFormatString_13);

                        var excelFormatString_14 = ws.Cells[14, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("#,##0.00\\ \"€\";[Red]\\-#,##0.00\\ \"€\"", excelFormatString_14);

                        var excelFormatString_15 = ws.Cells[15, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("0%", excelFormatString_15);

                        var excelFormatString_16 = ws.Cells[16, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("0.00%", excelFormatString_16);

                        var excelFormatString_17 = ws.Cells[17, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("0.00E+00", excelFormatString_17);

                        var excelFormatString_18 = ws.Cells[18, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("##0.0", excelFormatString_18);

                        var excelFormatString_19 = ws.Cells[19, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("# ?/?", excelFormatString_19);

                        var excelFormatString_20 = ws.Cells[20, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("# ??/??", excelFormatString_20);

                        var excelFormatString_21 = ws.Cells[21, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("mm-dd-yy", excelFormatString_21);

                        var excelFormatString_22 = ws.Cells[22, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("d-mmm-yy", excelFormatString_22);

                        var excelFormatString_23 = ws.Cells[23, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("d-mmm", excelFormatString_23);

                        var excelFormatString_24 = ws.Cells[24, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("mmm-yy", excelFormatString_24);

                        var excelFormatString_25 = ws.Cells[25, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("h:mm AM/PM", excelFormatString_25);

                        var excelFormatString_26 = ws.Cells[26, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("h:mm:ss AM/PM", excelFormatString_26);

                        var excelFormatString_27 = ws.Cells[27, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("h:mm", excelFormatString_27);

                        var excelFormatString_28 = ws.Cells[28, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("h:mm:ss", excelFormatString_28);

                        var excelFormatString_29 = ws.Cells[29, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("m/d/yy h:mm", excelFormatString_29);

                        var excelFormatString_30 = ws.Cells[30, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("mm:ss", excelFormatString_30);

                        var excelFormatString_31 = ws.Cells[31, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("mmss.0", excelFormatString_31);

                        var excelFormatString_32 = ws.Cells[32, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("@", excelFormatString_32);

                        var excelFormatString_33 = ws.Cells[33, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("[h]:mm:ss", excelFormatString_33);

                        var excelFormatString_34 = ws.Cells[34, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("_-* #,##0\\ \"€\"_-;\\-* #,##0\\ \"€\"_-;_-* \"-\"\\ \"€\"_-;_-@_-",
                            excelFormatString_34);

                        var excelFormatString_35 = ws.Cells[35, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("_-* #,##0\\ _€_-;\\-* #,##0\\ _€_-;_-* \"-\"\\ _€_-;_-@_-",
                            excelFormatString_35);

                        var excelFormatString_36 = ws.Cells[36, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("_-* #,##0.00\\ \"€\"_-;\\-* #,##0.00\\ \"€\"_-;_-* \"-\"??\\ \"€\"_-;_-@_-",
                            excelFormatString_36);

                        var excelFormatString_37 = ws.Cells[37, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("_-* #,##0.00\\ _€_-;\\-* #,##0.00\\ _€_-;_-* \"-\"??\\ _€_-;_-@_-",
                            excelFormatString_37);

                        var excelFormatString_38 = ws.Cells[38, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("mmm\\ yyyy", excelFormatString_38);

                        var excelFormatString_39 = ws.Cells[39, 1].Style?.Numberformat?.Format;
                        Assert.AreEqual("[$-407]dddd\\,\\ d/\\ mmmm\\ yyyy", excelFormatString_39);
                    }



                }
            }
        }



        [TestMethod]
        public void IssueWhitespaceInChartXml()
        {
            //Issue: If a Chart.xml contains ExtLst Nodes than the indentation of the chart.xml leads to corrupt Excefiles
            var excelTestFile = Resources.ChartIndentation;
            using (MemoryStream excelStream = new MemoryStream())
            {
                excelStream.Write(excelTestFile, 0, excelTestFile.Length);

                using (ExcelPackage exlPackage = new ExcelPackage(excelStream))
                {
                    var savePath = Path.Combine(TestContext.TestDeploymentDir, $"{TestContext.TestName}.xlsx");
                    exlPackage.SaveAs(new FileInfo(savePath));
                    var exApp = new Microsoft.Office.Interop.Excel.Application();

                    try
                    {
                        var exWbk = exApp.Workbooks.Open(savePath);
                    }
                    catch (System.Runtime.InteropServices.COMException)
                    {
                        Assert.Fail("It is not possible to open the workbook after EPPlus saved it.");
                    }
                    finally
                    {
                        exApp.Workbooks.Close();
                    }
                }
            }
        }


        [TestMethod]
        public void IssueTableWithXmlTags()
        {
            //Issue: If a cell is richtext and gets refrenced by another cell by formula the Cell gets the Xml-Node as Value.
            var excelTestFile = Resources.XMLTagsTable;
            using (MemoryStream excelStream = new MemoryStream())
            {
                excelStream.Write(excelTestFile, 0, excelTestFile.Length);


                using (ExcelPackage exlPackage = new ExcelPackage(excelStream))
                {
                    var sheet = exlPackage.Workbook.Worksheets["Tabelle1"];
                    Assert.AreEqual(sheet.Cells["A1"].Value, sheet.Cells["B1"].Value);
                    sheet.Calculate();
                    Assert.AreEqual(sheet.Cells["A1"].Value, sheet.Cells["B1"].Value);
                }
            }
        }

        [TestMethod]
        public void IssueWithVLookUpDateValue()
        {
            //Issue: If a VLookUp-Function contains a Date-Funktion as searchedValue an InvalidCastException is Thrown resulting in an #Value-Result
            var excelTestFile = Resources.VLookUpDateValue;
            using (MemoryStream excelStream = new MemoryStream())
            {
                excelStream.Write(excelTestFile, 0, excelTestFile.Length);

                using (ExcelPackage exlPackage = new ExcelPackage(excelStream))
                {
                    var ws = exlPackage.Workbook.Worksheets[1];

                    ws.Calculate();

                    Assert.AreEqual(ws.Cells["C2"].Value, ws.Cells["E3"].Value);

                }
            }
        }


        [TestMethod]
        public void IssueCanNotOpenAfterSaving()
        {
            //Issue: If a cell contains a hyperlink with special characters such as ä,ö,ü Excel encodes the link not in UTF-8 to keep the rule that a target link must be shorter than 2080 characters. 
            //Epplus always uses the normal UTF-8 encoding.In this case the Hyperlink would likely extend to over 2079 characters, resulting in a corrupt Excelfile that can not be opened from Excel.
            //To fix this problem I simply replaced the Hyperlink with a link that is a error - message, even though to fully mimic the behaviour from Excel a bigger fix would be necessary.

            var excelTestFile = Resources.HyperlinkIssue;

            using (MemoryStream excelStream = new MemoryStream())
            {
                excelStream.Write(excelTestFile, 0, excelTestFile.Length);

                using (ExcelPackage exlPackage = new ExcelPackage(excelStream))
                {

                    var savePath = Path.Combine(TestContext.TestDeploymentDir, $"{TestContext.TestName}.xlsx");
                    exlPackage.SaveAs(new FileInfo(savePath));
                    var exApp = new Microsoft.Office.Interop.Excel.Application();

                    try
                    {
                        var exWbk = exApp.Workbooks.Open(savePath);
                    }
                    catch (System.Runtime.InteropServices.COMException)
                    {
                        Assert.Fail("It is not possible to open the workbook after EPPlus saved it.");
                    }
                    finally
                    {
                        exApp.Workbooks.Close();
                    }
                }
            }
        }





        private TestContext _testContext;



        public TestContext TestContext
        {
            get => _testContext;
            set => _testContext = value;
        }
    public int Index
        {
            get { return _index; }
        }

        public void MoveIndexPointerForward()
        {
            _index++;
        }
    }
}