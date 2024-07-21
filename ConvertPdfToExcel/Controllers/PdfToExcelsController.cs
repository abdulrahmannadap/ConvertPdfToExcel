using ConvertPdfToExcel.DataBaseService;
using ConvertPdfToExcel.Models;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using UglyToad.PdfPig;

namespace ConvertPdfToExcel.Controllers
{
    public class PdfToExcelsController : Controller
    {
        #region Cunstructor_And_Depndancy
        private readonly ApplicationDbContext _context;

        public PdfToExcelsController(ApplicationDbContext context)
        {
            _context = context;
        }
        #endregion

        #region PdfToExcel_View_Page
        [HttpGet]
        public IActionResult Index()
        {
            return View();
        }
        #endregion

        #region  ConvertPdfToExcel_POST_Action
        [HttpPost]
        public IActionResult ConvertPdfToExcel(PdfToExcelModel model)
        {
            if (model.PdfFile == null || model.PdfFile.Length == 0)
            {
                return BadRequest("Please upload a valid PDF file.");
            }

            try
            {
                using (var memoryStream = new MemoryStream())
                {
                    model.PdfFile.CopyTo(memoryStream);
                    var pdfBytes = memoryStream.ToArray();

                    // Convert PDF bytes to Excel bytes and store in database 
                    var excelBytes = ConvertPdfToDynamicExcelBytes(pdfBytes);

                    // Create a filename
                    var excelFileName = Path.GetFileNameWithoutExtension(model.PdfFile.FileName) + ".xlsx";

                    // Store Excel data in database 
                    if (excelBytes != null)
                    {
                        var convertedData = new ConvertedExcelData
                        {
                            FileName = model.PdfFile.FileName,
                            ExcelData = excelBytes
                        };
                        _context.ConvertedExcelDatas.Add(convertedData);
                        _context.SaveChanges();
                    }

                    // Return the file for download (optional)
                    if (excelBytes != null)
                    {
                        return File(excelBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", excelFileName);
                    }
                    else
                    {
                        // Handle case where conversion fails (optional)
                        return BadRequest("Conversion failed. Please try again.");
                    }
                }
            }
            catch (Exception ex)
            {
                // Log the exception and handle database errors appropriately
                Console.WriteLine($"Error occurred while converting PDF: {ex.Message}");
                return BadRequest("An error occurred during processing. Please try again later.");
            }
        }
        #endregion

        //[HttpPost]
        //public IActionResult ConvertPdfToExcel(PdfToExcelModel model)
        //{
        //    if (model.PdfFile != null && model.PdfFile.Length > 0)
        //    {
        //        using (var memoryStream = new MemoryStream())
        //        {
        //            model.PdfFile.CopyTo(memoryStream);
        //            var pdfBytes = memoryStream.ToArray();

        //            // Convert PDF bytes to Excel bytes
        //            var excelBytes = ConvertPdfToDynamicExcelBytes(pdfBytes);

        //            // Create a filename
        //            var excelFileName = Path.GetFileNameWithoutExtension(model.PdfFile.FileName) + ".xlsx";

        //            // Return the file for download
        //            return File(excelBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", excelFileName);
        //        }
        //    }

        //    return View("Index");
        //}

        #region ConvertPdfToDynamicExcelBytes_Private_Methode
        private byte[] ConvertPdfToDynamicExcelBytes(byte[] pdfBytes)
        {
            using (var pdfDocument = PdfDocument.Open(new MemoryStream(pdfBytes)))
            {
                // Set the License Context for EPPlus
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("DynamicSheet");

                    int row = 1;
                    int column = 1;

                    // Track column names and indices
                    var columnNames = new Dictionary<string, int>();

                    // Loop through each page and line in the PDF document
                    foreach (var page in pdfDocument.GetPages())
                    {
                        var lines = page.Text.Split('\n');
                        foreach (var line in lines)
                        {
                            var parts = line.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);

                            // Assign column indices based on data structure
                            for (int i = 0; i < parts.Length; i++)
                            {
                                string header = $"Column {i + 1}"; // Generic column name
                                if (!columnNames.ContainsKey(header))
                                {
                                    columnNames[header] = column;
                                    worksheet.Cells[1, column].Value = header;
                                    worksheet.Cells[1, column].Style.Font.Bold = true;
                                    worksheet.Cells[1, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                    worksheet.Cells[1, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                                    column++;
                                }
                                worksheet.Cells[row + 1, columnNames[header]].Value = parts[i];
                            }
                            row++;
                        }
                    }

                    // Apply auto-filter to the columns
                    worksheet.Cells[worksheet.Dimension.Address].AutoFilter = true;

                    // Auto-fit columns
                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                    // Apply borders to the entire table for readability
                    var dataRange = worksheet.Cells[1, 1, row, columnNames.Count];
                    dataRange.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    dataRange.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    dataRange.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    dataRange.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                    return package.GetAsByteArray();
                }
            }
        }
        #endregion

        #region DownloadExcel_GET_Action
        [HttpGet]
        public IActionResult DownloadExcel()
        {
            return View();
        }
        #endregion

        #region DownloadExcel_POST_Action
        [HttpPost]
        public IActionResult DownloadExcel(int id)
        {
            try
            {
                var data = _context.ConvertedExcelDatas.FirstOrDefault(x => x.Id == id);
                if (data == null)
                {
                    return NotFound();
                }

                return File(data.ExcelData, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", data.FileName + ".xlsx");
            }
            catch (Exception ex)
            {
                // Log the exception and handle database errors appropriately
                Console.WriteLine($"Error occurred while retrieving data: {ex.Message}");
                return BadRequest("An error occurred. Please try again later.");
            }
        }
        #endregion

    }
}
