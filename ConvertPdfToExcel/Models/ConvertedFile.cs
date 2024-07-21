namespace ConvertPdfToExcel.Models
{
    public class ConvertedFile
    {
        public int Id { get; set; }
        public string FileName { get; set; }
        public byte[] FileContent { get; set; }
        public DateTime UploadedAt { get; set; }
    }
}
