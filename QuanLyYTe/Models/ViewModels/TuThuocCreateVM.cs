using System.ComponentModel.DataAnnotations;

namespace QuanLyYTe.Models.ViewModels
{
    public class TuThuocCreateVM
    {
        [Required]
        public string TenTuThuoc { get; set; } = null!;

        [Required]
        public int ID_PhongBan { get; set; }

        [Required]
        public decimal Latitude { get; set; }

        [Required]
        public decimal Longitude { get; set; }

        [Required]
        public string GhiChu { get; set; } = null!;
    }
}
