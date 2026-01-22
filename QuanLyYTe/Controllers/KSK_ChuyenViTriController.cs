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
    public class KSK_ChuyenViTriController : Controller
    {
        private readonly DataContext _context;
        private readonly IWebHostEnvironment _webHostEnvironment;
        private readonly IFileStorageService _fileStorageService;

        public KSK_ChuyenViTriController(DataContext _context, IWebHostEnvironment webHostEnvironment, IFileStorageService fileStorageService)
        {
            this._context = _context;
            _webHostEnvironment = webHostEnvironment;
            _fileStorageService = fileStorageService;
        }
        public async Task<IActionResult> Index(string search, int page = 1)
        {
            var res = await (from a in _context.KSK_ChuyenViTri
                             join nv in _context.NhanVien on a.ID_NV equals nv.ID_NV
                             join bp in _context.PhongBan on nv.ID_PhongBan equals bp.ID_PhongBan
                             join k in _context.KipLamViec on nv.ID_Kip equals k.ID_Kip into ulist3
                             from k in ulist3.DefaultIfEmpty()
                             join vt in _context.ViTriLamViec on a.ID_ViTri equals vt.ID_ViTri into ulist4
                             from vt in ulist4.DefaultIfEmpty()
                             join ld in _context.LyDoKhongDat on a.LyDoKhongDat equals ld.ID_LyDo into ulist5
                             from ld in ulist5.DefaultIfEmpty()
                             join vts in _context.ViTriLamViec on a.ID_ViTri_ChuyenDen equals vts.ID_ViTri into ulist6
                             from vts in ulist6.DefaultIfEmpty()
                             select new KSK_ChuyenViTri
                             {
                                 ID_KSK_CVT = a.ID_KSK_CVT,
                                 ID_NV = (int)a.ID_NV,
                                 MaNV = nv.MaNV,
                                 HoTen = nv.HoTen,
                                 NgaySinh = (DateTime?)nv.NgaySinh ?? default,
                                 TenPhongBan = bp.TenPhongBan,
                                 TenKip = k.TenKip,
                                 TenViTri = vt.TenViTri,
                                 TenViTriSauKhiChuyen = vts.TenViTri,
                                 NgayKham = (DateTime?)a.NgayKham ?? default,
                                 Dat = a.Dat,
                                 KhongDat = a.KhongDat,
                                 LyDoKhongDat = (int?)a.LyDoKhongDat??default,
                                 TenLyDoKhongDat = ld.TenLyDo,
                                 GhiChu = a.GhiChu
                             }).ToListAsync();
            if (search != null)
            {
                res = res.Where(x => x.MaNV.ToLower().Contains(search.ToLower()) || x.HoTen.ToLower().Contains(search.ToLower())).ToList();
            }
            const int pageSize = 20;
            if (page < 1)
            {
                page = 1;
            }
            int resCount = res.Count;
            var pager = new Pager(resCount, page, pageSize);
            int recSkip = (page - 1) * pageSize;

            var ordered = res.OrderByDescending(x => x.NgayKham);
            var data = ordered.Skip(recSkip).Take(pager.PageSize).ToList();

            this.ViewBag.Pager = pager;
            this.ViewBag.search = search;
            return View(data);  

        }


        public async Task<IActionResult> Deatail(int? ID_NV, int page = 1)
        {
            var res = await (from a in _context.KSK_ChuyenViTri.Where(x=>x.ID_NV == ID_NV)
                             join nv in _context.NhanVien on a.ID_NV equals nv.ID_NV
                             join bp in _context.PhongBan on nv.ID_PhongBan equals bp.ID_PhongBan
                             join k in _context.KipLamViec on nv.ID_Kip equals k.ID_Kip into ulist3
                             from k in ulist3.DefaultIfEmpty()
                             join vt in _context.ViTriLamViec on a.ID_ViTri equals vt.ID_ViTri into ulist4
                             from vt in ulist4.DefaultIfEmpty()
                             join ld in _context.LyDoKhongDat on a.LyDoKhongDat equals ld.ID_LyDo into ulist5
                             from ld in ulist5.DefaultIfEmpty()
                             select new KSK_ChuyenViTri
                             {
                                 ID_KSK_CVT = a.ID_KSK_CVT,
                                 ID_NV = (int)a.ID_NV,
                                 MaNV = nv.MaNV,
                                 HoTen = nv.HoTen,
                                 NgaySinh = (DateTime?)nv.NgaySinh ?? default,
                                 TenPhongBan = bp.TenPhongBan,
                                 TenKip = k.TenKip,
                                 TenViTri = vt.TenViTri,
                                 NgayKham = (DateTime?)a.NgayKham ?? default,
                                 Dat = a.Dat,
                                 KhongDat = a.KhongDat,
                                 LyDoKhongDat = (int?)a.LyDoKhongDat ?? default,
                                 TenLyDoKhongDat = ld.TenLyDo,
                                 GhiChu = a.GhiChu,
                                 FileKhamSucKhoePath = a.FileKhamSucKhoePath
                             }).ToListAsync();
            var id_nv = _context.NhanVien.Where(x => x.ID_NV == ID_NV).FirstOrDefault();
            if (id_nv != null)
            {
                ViewBag.ID_NV = id_nv.MaNV;
            }

            const int pageSize = 20;
            if (page < 1)
            {
                page = 1;
            }
            int resCount = res.Count;
            var pager = new Pager(resCount, page, pageSize);
            int recSkip = (page - 1) * pageSize;
            var data = res.Skip(recSkip).Take(pager.PageSize).ToList();
            this.ViewBag.Pager = pager;
            return View(data);

        }
        public FileResult TestDownloadPCF()
        {
            string path = "Form files/BM_KSK_ChuyenViTri.xlsx";
            HttpContext.Response.ContentType = "application/xlsx";
            string filePath = Path.Combine(_webHostEnvironment.ContentRootPath, path);

            if (!System.IO.File.Exists(filePath))
            {
                return null; // Xử lý lỗi nếu file không tồn tại
            }
            List<ViTriLamViec> vt = _context.ViTriLamViec.ToList();
             List<LyDoKhongDat> ld = _context.LyDoKhongDat.ToList();

            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(2);
                for (var i = 0; i < vt.Count; i++)  
                {
                    worksheet.Cell(i + 2, 6).Value = vt[i].TenViTri;
                }
                for (var i = 0; i < ld.Count; i++)
                {
                    worksheet.Cell(i + 2, 5).Value = ld[i].TenLyDo;
                }
                // Lưu lại file Excel
                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    stream.Seek(0, SeekOrigin.Begin);
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", path);
                }
            }
        }
        public async Task<IActionResult> ImportExcel()
        {
            return PartialView();
        }
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> ImportExcel(IFormFile file)
        {
            try
            {
                if (file == null || file.Length == 0)
                {
                    return RedirectToAction("Index", "KSK_ChuyenViTri");
                }


                // Create the Directory if it is not exist
                string webRootPath = System.IO.Path.Combine(Directory.GetCurrentDirectory(), "wwwroot");
                string dirPath = Path.Combine(webRootPath, "ReceivedReports");
                if (!Directory.Exists(dirPath))
                {
                    Directory.CreateDirectory(dirPath);
                }

                // MAke sure that only Excel file is used 
                string dataFileName = Path.GetFileName(DateTime.Now.ToString("yyyyMMddHHmm"));

                string extension = Path.GetExtension(dataFileName);

                string[] allowedExtsnions = new string[] { ".xls", ".xlsx" };
                // Make a Copy of the Posted File from the Received HTTP Request
                string saveToPath = Path.Combine(dirPath, dataFileName);

                using (FileStream stream = new FileStream(saveToPath, FileMode.Create))
                {
                    file.CopyTo(stream);
                }

                // USe this to handle Encodeing differences in .NET Core
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                // read the excel file
                IExcelDataReader reader = null;
                using (var stream = new FileStream(saveToPath, FileMode.Open))
                {
                    if (extension == ".xlsx")
                        reader = ExcelReaderFactory.CreateBinaryReader(stream);
                    else
                        reader = ExcelReaderFactory.CreateOpenXmlReader(stream);

                    DataSet ds = new DataSet();
                    ds = reader.AsDataSet();
                    reader.Close();
                    if (ds != null && ds.Tables.Count > 0)
                    {
                        System.Data.DataTable serviceDetails = ds.Tables[0];

                        for (int i = 7; i < serviceDetails.Rows.Count; i++)
                        {
                            string MNV = serviceDetails.Rows[i][1].ToString().Trim();
                            var check_nv = _context.NhanVien.Where(x => x.MaNV == MNV).FirstOrDefault();
                            if (check_nv == null)
                            {
                                TempData["msgSuccess"] = "<script>alert('Vui lòng cập nhật dữ liệu nhân viên: " + MNV + "');</script>";
                                return RedirectToAction("Index", "KSK_ChuyenViTri");
                            }
                            string ViTri = serviceDetails.Rows[i][3].ToString().Trim();
                            var check_vitri = _context.ViTriLamViec.Where(x => x.TenViTri == ViTri).FirstOrDefault();
                            string viTriChuyenDen = serviceDetails.Rows[i][4].ToString().Trim();
                            var check_viTriChuyenDen = _context.ViTriLamViec.Where(x => x.TenViTri == viTriChuyenDen).FirstOrDefault();
                            if (check_vitri == null)
                            {
                                TempData["msgSuccess"] = "<script>alert('Vui lòng kiểm tra vị trí. Nhân viên: " + MNV + "');</script>";

                                return RedirectToAction("Index", "KSK_ChuyenViTri");
                            }

                            if (check_viTriChuyenDen == null)
                            {
                                TempData["msgSuccess"] = "<script>alert('Vui lòng kiểm tra vị trí chuyển đến. Nhân viên: " + MNV + "');</script>";

                                return RedirectToAction("Index", "KSK_ChuyenViTri");
                            }

                            string Dat = serviceDetails.Rows[i][5].ToString().Trim();
                            string KhongDat = serviceDetails.Rows[i][6].ToString().Trim();
                            string LyDoKhongDat = serviceDetails.Rows[i][7].ToString().Trim();
                            var check_ld = _context.LyDoKhongDat.Where(x=>x.TenLyDo == LyDoKhongDat).FirstOrDefault();
                            if(check_ld ==  null && KhongDat != "")
                            {
                                TempData["msgSuccess"] = "<script>alert('Vui lòng kiểm tra lý do không đạt: " + check_nv.HoTen + "');</script>";
                                return RedirectToAction("Index", "KSK_ChuyenViTri");
                            }
                            
                            string Ngay_Kham = serviceDetails.Rows[i][8].ToString().Trim();
                            DateTime NgayKham = DateTime.ParseExact(Ngay_Kham, "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None);
                            string GhiChu = serviceDetails.Rows[i][9].ToString().Trim();
                            var check_ = _context.KSK_ChuyenViTri.Where(x => x.ID_NV == check_nv.ID_NV && x.NgayKham == NgayKham).FirstOrDefault();
                            if( check_ == null)
                            {
                                if (Dat != "")
                                {

                                    var result = _context.Database.ExecuteSqlRaw("EXEC KSK_ChuyenViTri_insert {0},{1},{2},{3},{4},{5},{6},{7}",
                                                                                  check_nv.ID_NV, check_vitri.ID_ViTri, NgayKham, Dat, KhongDat, null, GhiChu, check_viTriChuyenDen.ID_ViTri);
                                }
                                else
                                {

                                    var result = _context.Database.ExecuteSqlRaw("EXEC KSK_ChuyenViTri_insert {0},{1},{2},{3},{4},{5},{6},{7}",
                                                                                  check_nv.ID_NV, check_vitri.ID_ViTri, NgayKham, Dat, KhongDat, check_ld.ID_LyDo, GhiChu, check_viTriChuyenDen.ID_ViTri);
                                }
                            }    
                            else
                            {
                                if (Dat != "")
                                {

                                    var result = _context.Database.ExecuteSqlRaw("EXEC KSK_ChuyenViTri_update {0},{1},{2},{3},{4},{5},{6},{7}",
                                                                                 check_.ID_KSK_CVT, check_nv.ID_NV, check_vitri.ID_ViTri, NgayKham, Dat, KhongDat, null, GhiChu);
                                }
                                else
                                {

                                    var result = _context.Database.ExecuteSqlRaw("EXEC KSK_ChuyenViTri_update {0},{1},{2},{3},{4},{5},{6},{7}",
                                                                                  check_.ID_KSK_CVT, check_nv.ID_NV, check_vitri.ID_ViTri, NgayKham, Dat, KhongDat, check_ld.ID_LyDo, GhiChu);
                                }
                            }    
  
                        }
                    }
                }
                TempData["msgSuccess"] = "<script>alert('Import thành công');</script>";
            }
            catch (Exception e)
            {
                TempData["msgError"] = "<script>alert('Import thất bại');</script>";
            }

            return RedirectToAction("Index", "KSK_ChuyenViTri");
        }

        public async Task<IActionResult> Edit(int? id)
        {
            if (id == null)
            {
                TempData["msgError"] = "<script>alert('Chỉnh sửa thất bại');</script>";

                return RedirectToAction("Index", "KSK_ChuyenViTri");
            }

            var res = await (from a in _context.KSK_ChuyenViTri.Where(x=>x.ID_KSK_CVT == id)
                             join nv in _context.NhanVien on a.ID_NV equals nv.ID_NV
                             join bp in _context.PhongBan on nv.ID_PhongBan equals bp.ID_PhongBan
                             join k in _context.KipLamViec on nv.ID_Kip equals k.ID_Kip into ulist3
                             from k in ulist3.DefaultIfEmpty()
                             join vt in _context.ViTriLamViec on a.ID_ViTri equals vt.ID_ViTri into ulist4
                             from vt in ulist4.DefaultIfEmpty()
                             join vts in _context.ViTriLamViec on a.ID_ViTri_ChuyenDen equals vts.ID_ViTri into ulist5
                             from vts in ulist5.DefaultIfEmpty()
                             select new KSK_ChuyenViTri
                             {
                                 ID_KSK_CVT = a.ID_KSK_CVT,
                                 ID_NV = (int)a.ID_NV,
                                 MaNV = nv.MaNV,
                                 HoTen = nv.HoTen,
                                 NgaySinh = (DateTime?)nv.NgaySinh ?? default,
                                 TenPhongBan = bp.TenPhongBan,
                                 TenKip = k.TenKip,
                                 ID_ViTri = (int)a.ID_ViTri,
                                 TenViTri = vt.TenViTri,
                                 NgayKham = (DateTime?)a.NgayKham ?? default,
                                 Dat = a.Dat,
                                 KhongDat = a.KhongDat,
                                 LyDoKhongDat = a.LyDoKhongDat,
                                 GhiChu = a.GhiChu,
                                 ID_ViTri_ChuyenDen = a.ID_ViTri_ChuyenDen,
                                 TenViTriSauKhiChuyen = vts.TenViTri
                             }).ToListAsync();

            KSK_ChuyenViTri DO = new KSK_ChuyenViTri();
            if (res.Count > 0)
            {
                foreach (var a in res)
                {
                    DO.ID_KSK_CVT = a.ID_KSK_CVT;
                    DO.ID_NV = (int)a.ID_NV;
                    DO.ID_ViTri = (int)a.ID_ViTri;
                    DO.NgayKham = (DateTime?)a.NgayKham ?? default;
                    DO.Dat = a.Dat;
                    DO.KhongDat = a.KhongDat;
                    DO.LyDoKhongDat = a.LyDoKhongDat;
                    DO.GhiChu = a.GhiChu;
                    DO.ID_ViTri_ChuyenDen = a.ID_ViTri_ChuyenDen;
                }

                var NhanVien = (from nv in _context.NhanVien
                                select new NhanVien
                                {
                                    ID_NV = (int)nv.ID_NV,
                                    HoTen = nv.MaNV + " : " + nv.HoTen
                                }).ToList();
                ViewBag.NVList = new SelectList(NhanVien, "ID_NV", "HoTen", DO.ID_NV);


                List<ViTriLamViec> vt = _context.ViTriLamViec.ToList();
                ViewBag.VTList = new SelectList(vt, "ID_ViTri", "TenViTri", DO.ID_ViTri);

                List<LyDoKhongDat> ld = _context.LyDoKhongDat.ToList();
                ViewBag.LDList = new SelectList(ld, "ID_LyDo", "TenLyDo", DO.LyDoKhongDat);

                DateTime NK = (DateTime)DO.NgayKham;
                ViewBag.NgayKham = NK.ToString("yyyy-MM-dd");
            }
            else
            {
                return NotFound();
            }

            return PartialView(DO);
        }
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(int id, KSK_ChuyenViTri _DO)
        {
            try
            {
                if(_DO.LyDoKhongDat == 0 || _DO.LyDoKhongDat == null)
                {
                    var result = _context.Database.ExecuteSqlRaw("EXEC KSK_ChuyenViTri_update {0},{1},{2},{3},{4},{5},{6},{7}, {8}",
                                                                                       _DO.ID_KSK_CVT, _DO.ID_NV, _DO.ID_ViTri, _DO.NgayKham, _DO.Dat, _DO.KhongDat, null, _DO.GhiChu, _DO.ID_ViTri_ChuyenDen);
                }    
                else 
                {
                    var result = _context.Database.ExecuteSqlRaw("EXEC KSK_ChuyenViTri_update {0},{1},{2},{3},{4},{5},{6},{7}, {8}",
                                                                                       _DO.ID_KSK_CVT, _DO.ID_NV, _DO.ID_ViTri, _DO.NgayKham, _DO.Dat, _DO.KhongDat, _DO.LyDoKhongDat, _DO.GhiChu, _DO.ID_ViTri_ChuyenDen);
                }    

                TempData["msgSuccess"] = "<script>alert('Chỉnh sửa thành công');</script>";
            }
            catch (Exception e)
            {
                TempData["msgError"] = "<script>alert('Chỉnh sửa thất bại');</script>";
            }


            return RedirectToAction("Index", "KSK_ChuyenViTri");
        }

        public async Task<IActionResult> Delete(int id, int? page)
        {
            try
            {
                var result = _context.Database.ExecuteSqlRaw("EXEC KSK_ChuyenViTri_delete {0}", id);

                TempData["msgSuccess"] = "<script>alert('Xóa thành công');</script>";
            }
            catch (Exception e)
            {
                TempData["msgError"] = "<script>alert('Xóa dữ liệu thất bại');</script>";
            }


            return RedirectToAction("Index", "KSK_ChuyenViTri", new { page = page });
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> UploadFileKham(KSKChuyenViTriUploadVM model)
        {
            if (!ModelState.IsValid)
            {
                TempData["msgError"] = "<script>alert('Dữ liệu không hợp lệ');</script>";
                return RedirectToAction(nameof(Index));
            }

            var entity = await _context.KSK_ChuyenViTri
                .FirstOrDefaultAsync(x => x.ID_KSK_CVT == model.ID_KSK_CVT);

            if (entity == null)
            {
                TempData["msgError"] = "<script>alert('Không tìm thấy hồ sơ');</script>";
                return RedirectToAction(nameof(Index));
            }

            var uploadResult = await _fileStorageService.UploadAsync(
                model.FileKhamSucKhoe,
                folder: "uploads/ksk",
                maxSizeInBytes: 10 * 1024 * 1024 // 10MB
            );

            if (!uploadResult.Success)
            {
                TempData["msgError"] = uploadResult.ErrorMessage;
                return RedirectToAction(nameof(Index));
            }

            // Xóa file cũ SAU KHI upload file mới thành công
            if (!string.IsNullOrEmpty(entity.FileKhamSucKhoePath))
            {
                _fileStorageService.Delete(entity.FileKhamSucKhoePath);
            }

            entity.FileKhamSucKhoePath = uploadResult.FilePath;
            entity.FileKhamSucKhoeName = uploadResult.FileName;
            entity.FileKhamSucKhoeSize = uploadResult.FileSize;
            entity.FileKhamSucKhoeType = uploadResult.ContentType;

            await _context.SaveChangesAsync();

            TempData["msgSuccess"] = "<script>alert('Upload hồ sơ khám sức khỏe thành công');</script>";

            return RedirectToAction("Deatail", "KSK_ChuyenViTri", new { ID_NV = entity.ID_NV });
        }

        [HttpGet]
        public async Task<IActionResult> DownloadFileKham(int id)
        {
            var entity = await _context.KSK_ChuyenViTri
                .AsNoTracking()
                .FirstOrDefaultAsync(x => x.ID_KSK_CVT == id);

            if (entity == null || string.IsNullOrEmpty(entity.FileKhamSucKhoePath))
            {
                TempData["msgError"] = "<script>alert('File không tồn tại');</script>";
                return RedirectToAction(nameof(Index));
            }

            var fileResult = _fileStorageService.GetFile(entity.FileKhamSucKhoePath);

            if (!fileResult.Success)
            {
                TempData["msgError"] = fileResult.ErrorMessage;
                return RedirectToAction(nameof(Index));
            }

            return File(
                fileResult.FileBytes!,
                entity.FileKhamSucKhoeType ?? "application/octet-stream",
                entity.FileKhamSucKhoeName ?? "file"
            );
        }
    }
}
