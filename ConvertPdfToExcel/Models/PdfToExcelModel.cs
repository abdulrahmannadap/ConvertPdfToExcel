using System.ComponentModel.DataAnnotations;

namespace ConvertPdfToExcel.Models
{
    public class PdfToExcelModel
    {
        public IFormFile PdfFile { get; set; } 
    }
}
