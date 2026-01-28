namespace QuanLyYTe.Models.ViewModels
{
    public class FileDownloadResult
    {
        public bool Success { get; private set; }
        public byte[]? FileBytes { get; private set; }
        public string? ErrorMessage { get; private set; }

        public static FileDownloadResult Fail(string message)
        {
            return new FileDownloadResult
            {
                Success = false,
                ErrorMessage = message
            };
        }

        public static FileDownloadResult Ok(byte[] bytes)
        {
            return new FileDownloadResult
            {
                Success = true,
                FileBytes = bytes
            };
        }
    }
}
