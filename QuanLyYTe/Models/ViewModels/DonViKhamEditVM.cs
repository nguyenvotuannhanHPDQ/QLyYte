using System.ComponentModel.DataAnnotations;

namespace QuanLyYTe.Models.ViewModels
{
    public class DonViKhamEditVM
    {
        public int ID_DonViKham { get; set; }

        [Required]
        public string TenDonVi { get; set; } = string.Empty;

        public List<IFormFile>? Files { get; set; }

        public List<HoSoFileVM> ExistingFiles { get; set; } = new();
    }
}
