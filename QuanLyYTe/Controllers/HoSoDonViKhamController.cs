using ClosedXML.Excel;
using ExcelDataReader;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using QuanLyYTe.Models;
using QuanLyYTe.Models.ViewModels;
using QuanLyYTe.Repositorys;
using QuanLyYTe.Services.Interfaces;
using System.Data;

namespace QuanLyYTe.Controllers
{
    public class HoSoDonViKhamController : Controller
    {
        private readonly DataContext _context;
        private readonly IWebHostEnvironment _webHostEnvironment;
        private readonly IFileStorageService _fileStorageService;

        public HoSoDonViKhamController(DataContext context, IWebHostEnvironment webHostEnvironment, IFileStorageService fileStorageService)
        {
            _context = context;
            _webHostEnvironment = webHostEnvironment;
            _fileStorageService = fileStorageService;
        }

        public IActionResult Index(string? search)
        {
            var query = _context.DM_DonViKham
                .AsNoTracking()
                .AsQueryable();

            if (!string.IsNullOrWhiteSpace(search))
            {
                var keyword = search.Trim();
                query = query.Where(x => x.TenDonVi.Contains(keyword));
            }

            var data = query
                .Select(dv => new DonViKhamHoSoVM
                {
                    ID_DonViKham = dv.ID_DonViKham,
                    TenDonVi = dv.TenDonVi,

                    HoSoFiles = dv.KSK_HoSoDonVi
                        .Where(f => !string.IsNullOrEmpty(f.FilePath))
                        .Select(f => new HoSoFileVM
                        {
                            ID_HoSo = f.ID_HoSo,
                            TenFile = f.TenFile,
                            FilePath = f.FilePath
                        })
                        .ToList()
                })
                .OrderBy(x => x.TenDonVi)
                .ToList();

            ViewBag.search = search;

            return View(data);
        }

        public IActionResult Create()
        {
            return View();
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create(string TenDonVi, List<IFormFile> Files)
        {
            if (string.IsNullOrWhiteSpace(TenDonVi))
            {
                ModelState.AddModelError("", "Tên đơn vị không được để trống.");
                return View();
            }

            if (Files == null || !Files.Any(f => f != null && f.Length > 0))
            {
                ModelState.AddModelError("", "Vui lòng chọn ít nhất một file hồ sơ.");
                return View();
            }

            using var transaction = await _context.Database.BeginTransactionAsync();

            try
            {
                var donVi = new DM_DonViKham
                {
                    TenDonVi = TenDonVi.Trim(),
                    IsActive = true,
                    CreatedDate = DateTimeSafe.Now()
                };

                _context.DM_DonViKham.Add(donVi);
                await _context.SaveChangesAsync();

                var uploadPath = Path.Combine(
                    Directory.GetCurrentDirectory(),
                    "wwwroot",
                    "Uploads",
                    "DonViKham"
                );

                if (!Directory.Exists(uploadPath))
                    Directory.CreateDirectory(uploadPath);

                foreach (var file in Files.Where(f => f != null && f.Length > 0))
                {
                    var originalName = Path.GetFileName(file.FileName);
                    var savedName = $"{Guid.NewGuid()}_{originalName}";
                    var fullPath = Path.Combine(uploadPath, savedName);

                    using (var stream = new FileStream(fullPath, FileMode.Create))
                    {
                        await file.CopyToAsync(stream);
                    }

                    _context.KSK_HoSoDonVi.Add(new KSK_HoSoDonVi
                    {
                        ID_DonViKham = donVi.ID_DonViKham,
                        TenHoSo = Path.GetFileNameWithoutExtension(originalName),
                        TenFile = originalName,
                        FilePath = "/Uploads/DonViKham/" + savedName,
                        FileType = Path.GetExtension(originalName),
                        FileSize = file.Length,
                        NgayUpload = DateTimeSafe.Now(),
                        IsActive = true
                    });
                }

                await _context.SaveChangesAsync();
                await transaction.CommitAsync();

                TempData["msgSuccess"] = "<script>alert('Thêm mới thành công');</script>";

                return RedirectToAction("Index");
            }
            catch
            {
                await transaction.RollbackAsync();
                ModelState.AddModelError("", "Có lỗi xảy ra khi lưu dữ liệu.");
                return View();
            }
        }

        public IActionResult DownloadFile(int id)
        {
            var file = _context.KSK_HoSoDonVi
                .AsNoTracking()
                .FirstOrDefault(x => x.ID_HoSo == id);

            if (file == null || string.IsNullOrEmpty(file.FilePath))
                return NotFound();

            var fullPath = Path.Combine(_webHostEnvironment.WebRootPath, file.FilePath.TrimStart('/'));

            if (!System.IO.File.Exists(fullPath))
                return NotFound();

            var contentType = "application/octet-stream";
            var fileName = file.TenFile ?? Path.GetFileName(fullPath);

            return PhysicalFile(fullPath, contentType, fileName);
        }

        [HttpGet]
        public IActionResult Edit(int id)
        {
            var entity = _context.DM_DonViKham
                .AsNoTracking()
                .Include(x => x.KSK_HoSoDonVi)
                .FirstOrDefault(x => x.ID_DonViKham == id);

            if (entity == null)
                return NotFound();

            var vm = new DonViKhamEditVM
            {
                ID_DonViKham = entity.ID_DonViKham,
                TenDonVi = entity.TenDonVi,

                ExistingFiles = entity.KSK_HoSoDonVi
                    .Select(f => new HoSoFileVM
                    {
                        ID_HoSo = f.ID_HoSo,
                        TenFile = f.TenFile,
                        FilePath = f.FilePath
                    })
                    .ToList()
            };

            return View(vm);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(DonViKhamEditVM model)
        {
            if (!ModelState.IsValid)
                return View(model);

            var entity = await _context.DM_DonViKham
                .Include(x => x.KSK_HoSoDonVi)
                .FirstOrDefaultAsync(x => x.ID_DonViKham == model.ID_DonViKham);

            if (entity == null)
                return NotFound();

            entity.TenDonVi = model.TenDonVi.Trim();
            entity.CreatedDate = DateTimeSafe.Now();

            if (model.Files != null && model.Files.Any())
            {
                foreach (var file in model.Files.Where(f => f.Length > 0))
                {
                    var fileName = Path.GetFileName(file.FileName);
                    var savePath = Path.Combine("wwwroot/Uploads/DonViKham", fileName);

                    using (var stream = new FileStream(savePath, FileMode.Create))
                    {
                        await file.CopyToAsync(stream);
                    }

                    entity.KSK_HoSoDonVi.Add(new KSK_HoSoDonVi
                    {
                        TenFile = fileName,
                        FilePath = "/Uploads/DonViKham/" + fileName,
                        NgayUpload = DateTimeSafe.Now(),
                        TenHoSo = Path.GetFileNameWithoutExtension(fileName),
                        FileType = Path.GetExtension(fileName),
                        FileSize = file.Length,
                        IsActive = true
                    });
                }
            }

            await _context.SaveChangesAsync();
            TempData["msgSuccess"] = "<script>alert('Cập nhật thành công');</script>";

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
