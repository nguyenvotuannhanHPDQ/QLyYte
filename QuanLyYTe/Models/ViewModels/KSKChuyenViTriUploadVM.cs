using System.ComponentModel.DataAnnotations;

namespace QuanLyYTe.Models.ViewModels
{
    public class KSKChuyenViTriUploadVM
    {
        [Required]
        public int ID_KSK_CVT { get; set; }

        [Required]
        public IFormFile? FileKhamSucKhoe { get; set; }
    }

}
