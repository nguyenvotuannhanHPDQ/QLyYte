using QuanLyYTe.Models.ViewModels;

namespace QuanLyYTe.Services.Interfaces
{
    public interface IFileStorageService
    {
        Task<FileUploadResult> UploadAsync(
            IFormFile file,
            string folder,
            long maxSizeInBytes);

        bool Delete(string filePath);

        FileDownloadResult GetFile(string filePath);
    }
}
