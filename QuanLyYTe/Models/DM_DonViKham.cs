using System.ComponentModel.DataAnnotations;

namespace QuanLyYTe.Models
{
    public class DM_DonViKham
    {
        [Key]
        public int ID_DonViKham { get; set; }

        public string? MaDonVi { get; set; }
        public string TenDonVi { get; set; } = null!;
        public string? DiaChi { get; set; }
        public string? DienThoai { get; set; }
        public string? Email { get; set; }
        public string? GhiChu { get; set; }
        public bool IsActive { get; set; }
        public DateTime CreatedDate { get; set; }

        public virtual ICollection<KSK_HoSoDonVi> KSK_HoSoDonVi { get; set; }
            = new List<KSK_HoSoDonVi>();
    }

}
