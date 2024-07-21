//using ConvertPdfToExcel.Models;
//using iTextSharp.text.pdf;
//using iTextSharp.text.pdf.parser;
//using Microsoft.AspNetCore.Mvc;
//using OfficeOpenXml;
//using System.Text;

//namespace PdfToExcel.Controllers
//{
//    public class PdfToExcelssController : Controller
//    {
//        [HttpGet]
//        public IActionResult Index()
//        {
//            return View();
//        }

//        [HttpPost]
//        public IActionResult ConvertPdfToExcel(PdfToExcelModels pdfToExcelModels)
//        {
//            if (pdfToExcelModels == null || pdfToExcelModels.Length == 0)
//            {
//                return BadRequest("No file uploaded");
//            }

//            using (MemoryStream pdfStream = new MemoryStream())
//            {
//                pdfToExcelModels.CopyTo(pdfStream);
//                pdfStream.Position = 0;

//                using (PdfReader pdfReader = new PdfReader(pdfStream))
//                {
//                    StringBuilder sb = new StringBuilder();
//                    for (int page = 1; page <= pdfReader.NumberOfPages; page++)
//                    {
//                        ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
//                        string currentPageText = PdfTextExtractor.GetTextFromPage(pdfReader, page, strategy);
//                        sb.Append(currentPageText);
//                    }

//                    // Split the text into rows and columns
//                    string[] rows = sb.ToString().Split('\n');
//                    string[][] data = new string[rows.Length][];
//                    for (int i = 0; i < rows.Length; i++)
//                    {
//                        data[i] = rows[i].Split('\t');
//                    }

//                    using (MemoryStream excelStream = new MemoryStream())
//                    {
//                        using (ExcelPackage package = new ExcelPackage(excelStream))
//                        {
//                            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");

//                            // Add the data to the Excel worksheet
//                            for (int i = 0; i < data.Length; i++)
//                            {
//                                for (int j = 0; j < data[i].Length; j++)
//                                {
//                                    worksheet.Cells[i + 1, j + 1].Value = data[i][j];
//                                }
//                            }

//                            package.Save();
//                        }

//                        // Return the Excel file as a download
//                        return File(excelStream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "output.xlsx");
//                    }
//                }
//            }
//        }
//    }
//}