using ConvertPdfToExcel.DataBaseService;
using ConvertPdfToExcel.Models;
using Microsoft.AspNetCore.Mvc;
using Spire.Pdf;
using Spire.Xls;

namespace ConvertPdfToExcel.Controllers
{
    public class PdfToExcelController : Controller
    {
        private readonly ApplicationDbContext _context;

        public PdfToExcelController(ApplicationDbContext context)
        {
            _context = context;
        }

        [HttpGet]
        public IActionResult Index()
        {
            return View();
        }
        #region ExportToExcel_Action
        [HttpPost]
        public async Task<IActionResult> ExportToExcel(PdfInputModel model)
        {
            if (model.PdfFile == null || model.PdfFile.Length == 0)
            {
                ModelState.AddModelError("PdfFile", "Please upload a valid PDF file.");
                return View("Index");
            }

            // Extract the base name of the uploaded file
            string originalFileName = Path.GetFileNameWithoutExtension(model.PdfFile.FileName);
            string sanitizedFileName = SanitizeFileName(originalFileName);
            string outputFileName = $"{sanitizedFileName}.xlsx";

            using var pdfStream = new MemoryStream();
            model.PdfFile.CopyTo(pdfStream);

            // Load the PDF document
            PdfDocument pdfDocument = new PdfDocument();
            pdfDocument.LoadFromStream(pdfStream);

            // Create a new workbook and worksheet using Spire.Xls
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            int rowIndex = 1;

            // Iterate through the pages of the PDF document
            for (int i = 0; i < pdfDocument.Pages.Count; i++)
            {
                PdfPageBase page = pdfDocument.Pages[i];
                // Extract text from the page
                string text = page.ExtractText();

                // Write the text to the Excel sheet
                sheet.Range[$"A{rowIndex}"].Text = text;
                rowIndex += 1; // Move to the next row for the next page's text
            }

            // Save workbook to memory stream
            using MemoryStream excelStream = new MemoryStream();
            workbook.SaveToStream(excelStream, Spire.Xls.FileFormat.Version2013);

            // Save the file to the database
            var convertedFile = new ConvertedFile
            {
                FileName = outputFileName,
                FileContent = excelStream.ToArray(),
                UploadedAt = DateTime.UtcNow
            };

            _context.ConvertedFiles.Add(convertedFile);
            await _context.SaveChangesAsync();

            // Return the file to the user for auto-download
            return File(convertedFile.FileContent,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        convertedFile.FileName);
        }

        // Utility method to sanitize file names
        private string SanitizeFileName(string fileName)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                fileName = fileName.Replace(c, '_');
            }
            return fileName;
        }
        #endregion
    }
}
