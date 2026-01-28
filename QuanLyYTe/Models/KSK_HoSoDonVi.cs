using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace QuanLyYTe.Models
{
    public class KSK_HoSoDonVi
    {
        [Key]
        public int ID_HoSo { get; set; }

        public int ID_DonViKham { get; set; }

        public string TenHoSo { get; set; } = null!;
        public string TenFile { get; set; } = null!;
        public string FilePath { get; set; } = null!;
        public long? FileSize { get; set; }
        public string? FileType { get; set; }
        public DateTime NgayUpload { get; set; }
        public string? NguoiUpload { get; set; }
        public string? GhiChu { get; set; }
        public bool IsActive { get; set; }

        [ForeignKey(nameof(ID_DonViKham))]
        public virtual DM_DonViKham DM_DonViKham { get; set; } = null!;
    }

}
