namespace ConvertPdfToExcel.Models
{
    public class ConvertedExcelData
    {
        public int Id { get; set; } // Primary key for the database table
        public string FileName { get; set; } // Name of the original PDF file
        public byte[] ExcelData { get; set; } // The converted Excel file content
                                             
    }
}
