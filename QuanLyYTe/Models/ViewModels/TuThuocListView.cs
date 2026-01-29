namespace QuanLyYTe.Models.ViewModels
{
    public class TuThuocListView
    {
        public int ID_TuThuoc { get; set; }
        public string TenTuThuoc { get; set; }
        public string TenPhongBan { get; set; }
        public string GhiChu { get; set; }
        public decimal Latitude { get; set; }
        public decimal Longitude { get; set; }
        public DateTime NgayTao { get; set; }

        public string MapUrl =>
            $"https://www.openstreetmap.org/?mlat={Latitude}&mlon={Longitude}#map=18/{Latitude}/{Longitude}";
    }
}
