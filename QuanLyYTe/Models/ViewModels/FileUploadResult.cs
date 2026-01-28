namespace QuanLyYTe.Models.ViewModels
{
    public class FileUploadResult
    {
        public bool Success { get; private set; }

        public string? FilePath { get; private set; }
        public string? FileName { get; private set; }
        public long FileSize { get; private set; }
        public string? ContentType { get; private set; }

        public string? ErrorMessage { get; private set; }

        public static FileUploadResult Fail(string message)
        {
            return new FileUploadResult
            {
                Success = false,
                ErrorMessage = message
            };
        }

        public static FileUploadResult Ok(
            string filePath,
            string fileName,
            long fileSize,
            string contentType)
        {
            return new FileUploadResult
            {
                Success = true,
                FilePath = filePath,
                FileName = fileName,
                FileSize = fileSize,
                ContentType = contentType
            };
        }
    }

}
