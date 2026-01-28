using System.ComponentModel.DataAnnotations;

namespace QuanLyYTe.Models.ViewModels
{
    public class KSKSoCapCuuUploadVM
    {
        [Required]
        public int ID_SCC { get; set; }

        [Required]
        public IFormFile? FileKhamSucKhoe { get; set; }
    }
}
