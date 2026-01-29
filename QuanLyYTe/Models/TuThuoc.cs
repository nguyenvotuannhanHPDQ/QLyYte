using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace QuanLyYTe.Models
{
    public class TuThuoc
    {
        [Key]
        public int ID_TuThuoc { get; set; }
        public string TenTuThuoc { get; set; } = null!;
        public int ID_PhongBan { get; set; }

        public decimal Latitude { get; set; }
        public decimal Longitude { get; set; }

        public string GhiChu { get; set; }
        public bool IsActive { get; set; }
        public DateTime CreatedAt { get; set; }

        [ForeignKey(nameof(ID_PhongBan))]
        public PhongBan? PhongBan { get; set; }
    }
}
