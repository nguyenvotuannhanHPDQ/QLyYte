using QuanLyYTe.Models.ViewModels;
using QuanLyYTe.Services.Interfaces;

namespace QuanLyYTe.Services.Implementions
{
    public class FileStorageService : IFileStorageService
    {
        private readonly IWebHostEnvironment _env;

        private static readonly string[] AllowedContentTypes =
        {
            "application/pdf",
            "image/jpeg",
            "image/png"
        };

        public FileStorageService(IWebHostEnvironment env)
        {
            _env = env;
        }

        public async Task<FileUploadResult> UploadAsync(
            IFormFile file,
            string folder,
            long maxSizeInBytes)
        {
            if (file == null || file.Length == 0)
                return FileUploadResult.Fail("File rỗng hoặc không hợp lệ");

            if (file.Length > maxSizeInBytes)
                return FileUploadResult.Fail("Dung lượng file vượt quá giới hạn");

            if (!AllowedContentTypes.Contains(file.ContentType))
                return FileUploadResult.Fail("Định dạng file không được phép");

            try
            {
                var uploadFolder = Path.Combine(_env.WebRootPath, folder);
                Directory.CreateDirectory(uploadFolder);

                var extension = Path.GetExtension(file.FileName);
                var storedFileName = $"{Guid.NewGuid()}{extension}";
                var fullPath = Path.Combine(uploadFolder, storedFileName);

                using var stream = new FileStream(fullPath, FileMode.Create);
                await file.CopyToAsync(stream);

                return FileUploadResult.Ok(
                    filePath: $"/{folder}/{storedFileName}",
                    fileName: file.FileName,
                    fileSize: file.Length,
                    contentType: file.ContentType
                );
            }
            catch
            {
                // Không throw – trả kết quả
                return FileUploadResult.Fail("Lỗi khi lưu file");
            }
        }

        public bool Delete(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                return false;

            try
            {
                var fullPath = Path.Combine(
                    _env.WebRootPath,
                    filePath.TrimStart('/')
                );

                if (!File.Exists(fullPath))
                    return false;

                File.Delete(fullPath);
                return true;
            }
            catch
            {
                return false;
            }
        }

        public FileDownloadResult GetFile(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                return FileDownloadResult.Fail("Đường dẫn file không hợp lệ");

            try
            {
                var fullPath = Path.Combine(
                    _env.WebRootPath,
                    filePath.TrimStart('/')
                );

                if (!System.IO.File.Exists(fullPath))
                    return FileDownloadResult.Fail("File không tồn tại trên hệ thống");

                var bytes = System.IO.File.ReadAllBytes(fullPath);
                return FileDownloadResult.Ok(bytes);
            }
            catch
            {
                return FileDownloadResult.Fail("Không thể đọc file");
            }
        }

    }

}
