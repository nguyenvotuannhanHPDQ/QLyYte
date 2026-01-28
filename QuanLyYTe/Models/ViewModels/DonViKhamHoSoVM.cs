namespace QuanLyYTe.Models.ViewModels
{
    public class DonViKhamHoSoVM
    {
        public int ID_DonViKham { get; set; }
        public string TenDonVi { get; set; } = string.Empty;

        public List<HoSoFileVM> HoSoFiles { get; set; } = new();
    }

    public class HoSoFileVM
    {
        public int ID_HoSo { get; set; }
        public string TenFile { get; set; } = string.Empty;
        public string FilePath { get; set; } = string.Empty;
    }

}
