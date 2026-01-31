using DocumentFormat.OpenXml.InkML;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using QuanLyYTe.Models;
using QuanLyYTe.Models.ViewModels;
using QuanLyYTe.Repositorys;
using QuanLyYTe.Services.Interfaces;
using System.Security.Claims;
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

        [HttpGet]
        public IActionResult Delete(int id)
        {
            var TenDangNhap = User.FindFirstValue(ClaimTypes.Name);
            var list = _context.TaiKhoan.Where(x => x.TenDangNhap == TenDangNhap).FirstOrDefault();

            if (list != null && list.ID_Quyen != 1 && list.ID_Quyen != 2)
            {
                TempData["msgError"] = "<script>alert('Bạn không có quyền thực hiện chức năng này');</script>";
                return RedirectToAction(nameof(Index));
            }

            var entity = _context.TuThuoc
                .FirstOrDefault(x => x.ID_TuThuoc == id);

            if (entity == null)
            {
                TempData["msgError"] = "<script>alert('Không tìm thấy tủ thuốc cần xóa');</script>";
                return RedirectToAction(nameof(Index));
            }

            try
            {
                _context.TuThuoc.Remove(entity);
                _context.SaveChanges();

                TempData["msgSuccess"] = "<script>alert('Xóa tủ thuốc thành công');</script>";
            }
            catch (Exception)
            {
                TempData["msgError"] =
                    "<script>alert('Có lỗi xảy ra khi xóa dữ liệu');</script>";
            }

            return RedirectToAction(nameof(Index));
        }

        [HttpGet]
        public IActionResult Edit(int id)
        {
            var TenDangNhap = User.FindFirstValue(ClaimTypes.Name);
            var list = _context.TaiKhoan.Where(x => x.TenDangNhap == TenDangNhap).FirstOrDefault();

            if (list != null && list.ID_Quyen != 1 && list.ID_Quyen != 2)
            {
                TempData["msgError"] = "<script>alert('Bạn không có quyền thực hiện chức năng này');</script>";
                return RedirectToAction(nameof(Index));
            }

            ViewBag.BoPhan = _context.PhongBan
                .AsNoTracking()
                .OrderBy(x => x.TenPhongBan)
                .ToList();

            var data = _context.TuThuoc
                .AsNoTracking()
                .FirstOrDefault(x => x.ID_TuThuoc == id);

            if (data == null)
            {
                TempData["msgError"] = "<script>alert('Không tìm thấy dữ liệu');</script>";
                return RedirectToAction(nameof(Index));
            }

            ViewBag.BoPhan = _context.PhongBan
                .AsNoTracking()
                .OrderBy(x => x.TenPhongBan)
                .ToList();

            return View(data);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult Edit(int id, TuThuoc model, string LocationMode)
        {
            if (id <= 0)
            {
                TempData["msgError"] = "<script>alert('ID không hợp lệ');</script>";
                return RedirectToAction(nameof(Index));
            }

            if (string.IsNullOrWhiteSpace(model.TenTuThuoc)
                || model.ID_PhongBan <= 0
                || string.IsNullOrWhiteSpace(model.GhiChu))
            {
                TempData["msgError"] = "<script>alert('Vui lòng nhập đầy đủ thông tin');</script>";
                return RedirectToAction(nameof(Edit));
            }

            if (LocationMode == "map")
            {
                if (model.Latitude == null || model.Longitude == null)
                {
                    TempData["msgError"] = "<script>alert('Vui lòng chọn vị trí trên bản đồ');</script>";
                    return RedirectToAction(nameof(Edit));
                }
            }
            else if (LocationMode == "manual")
            {
                if (model.Latitude == null || model.Longitude == null)
                {
                    TempData["msgError"] = "<script>alert('Vui lòng nhập tọa độ hợp lệ');</script>";
                    return RedirectToAction(nameof(Edit));
                }
            }

            var entity = _context.TuThuoc.FirstOrDefault(x => x.ID_TuThuoc == id);
            if (entity == null)
            {
                TempData["msgError"] = "<script>alert('Không tìm thấy dữ liệu cần cập nhật');</script>";
                return RedirectToAction(nameof(Index));
            }

            try
            {
                entity.TenTuThuoc = model.TenTuThuoc.Trim();
                entity.ID_PhongBan = model.ID_PhongBan;
                entity.Latitude = model.Latitude;
                entity.Longitude = model.Longitude;
                entity.GhiChu = model.GhiChu;

                _context.SaveChanges();

                TempData["msgSuccess"] = "<script>alert('Cập nhật tủ thuốc thành công');</script>";
            }
            catch (Exception)
            {
                TempData["msgError"] = "<script>alert('Có lỗi xảy ra khi cập nhật');</script>";
            }

            return RedirectToAction(nameof(Index));
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
