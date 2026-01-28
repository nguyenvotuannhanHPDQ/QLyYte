using System.ComponentModel.DataAnnotations;

namespace QuanLyYTe.Models.ViewModels
{
    public class KSKDinhKyUploadVM
    {
        [Required]
        public int ID_KSK_DK { get; set; }

        [Required]
        public IFormFile? FileKhamSucKhoe { get; set; }
    }
}
