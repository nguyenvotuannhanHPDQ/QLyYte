using DocumentFormat.OpenXml.InkML;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using QuanLyYTe.Models;
using QuanLyYTe.Models.ViewModels;
using QuanLyYTe.Repositorys;
using QuanLyYTe.Services.Interfaces;
using static QuanLyYTe.Controllers.HoSoDonViKhamController;

namespace QuanLyYTe.Controllers
{
    public class TuThuocController : Controller
    {
        private readonly DataContext _context;
        private readonly IWebHostEnvironment _webHostEnvironment;
        private readonly IFileStorageService _fileStorageService;

        public TuThuocController(DataContext context, IWebHostEnvironment webHostEnvironment, IFileStorageService fileStorageService)
        {
            _context = context;
            _webHostEnvironment = webHostEnvironment;
            _fileStorageService = fileStorageService;
        }

        public IActionResult Index()
        {
            var data = _context.TuThuoc
                .Where(x => x.IsActive)
                .Select(x => new TuThuocListView
                {
                    ID_TuThuoc = x.ID_TuThuoc,
                    TenTuThuoc = x.TenTuThuoc,
                    TenPhongBan = x.PhongBan.TenPhongBan,
                    GhiChu = x.GhiChu,
                    Latitude = x.Latitude,
                    Longitude = x.Longitude,
                    NgayTao = x.CreatedAt
                })
                .OrderByDescending(x => x.ID_TuThuoc)
                .ToList();
            return View(data);
        }

        public IActionResult Create()
        {
            ViewBag.BoPhan = _context.PhongBan
                .AsNoTracking()
                .OrderBy(x => x.TenPhongBan)
                .ToList();

            return View();
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult Create(TuThuocCreateVM model)
        {
            if (!ModelState.IsValid)
            {
                TempData["msgError"] = "<script>alert('Vui lòng nhập đầy đủ thông tin và chọn vị trí trên bản đồ');</script>";
                return RedirectToAction(nameof(Create));
            }

            if (model.Latitude == 0 || model.Longitude == 0)
            {
                TempData["msgError"] =
                    "<script>alert('Vui lòng chọn vị trí trên bản đồ hoặc nhập tọa độ hợp lệ');</script>";
                return RedirectToAction(nameof(Create));
            }

            var entity = new TuThuoc

            {
                TenTuThuoc = model.TenTuThuoc.Trim(),
                ID_PhongBan = model.ID_PhongBan,
                Latitude = model.Latitude,
                Longitude = model.Longitude,
                GhiChu = model.GhiChu,
                IsActive = true,
                CreatedAt = DateTimeSafe.Now()
            };

            _context.TuThuoc.Add(entity);
            _context.SaveChanges();

            TempData["msgSuccess"] = "<script>alert('Thêm vị trí tủ thuốc thành công');</script>";
            return RedirectToAction(nameof(Index));
        }

        public  List<TuThuocListView> GetDanhSachTuThuoc()
        {
            return _context.TuThuoc
                .Where(x => x.IsActive)
                .Select(x => new TuThuocListView
                {
                    ID_TuThuoc = x.ID_TuThuoc,
                    TenTuThuoc = x.TenTuThuoc,
                    TenPhongBan = x.PhongBan.TenPhongBan,
                    GhiChu = x.GhiChu,
                    Latitude = x.Latitude,
                    Longitude = x.Longitude
                })
                .ToList();
        }

        public static class DateTimeSafe
        {
            private static readonly DateTime SqlMinDate = new DateTime(1753, 1, 1);

            public static DateTime Now()
            {
                return DateTime.Now < SqlMinDate
                    ? SqlMinDate
                    : DateTime.Now;
            }
        }
    }
}
