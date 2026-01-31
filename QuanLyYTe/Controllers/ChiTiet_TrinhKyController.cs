using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using QuanLyYTe.Models;
using QuanLyYTe.Repositorys;
using Microsoft.AspNetCore.Mvc.Rendering;
using ExcelDataReader;
using System.Data;
using Microsoft.Data.SqlClient;
using ClosedXML.Excel;
namespace QuanLyYTe.Controllers
{
    public class ChiTiet_TrinhKyController : Controller
    {
        private readonly DataContext _context;

        public ChiTiet_TrinhKyController(DataContext _context)
        {
            this._context = _context;
        }
        public async Task<IActionResult> Index(int id, int page = 1)
        
        {
            var check = _context.TrinhKy.Where(x=>x.ID_TK == id).FirstOrDefault();

            var res = await (from a in _context.KSK_BenhNgheNghiep.Where(x => x.ID_PhongBan == check.ID_PhongBan)
                             join nv in _context.NhanVien on a.ID_NV equals nv.ID_NV
                             join bp in _context.PhongBan on a.ID_PhongBan equals bp.ID_PhongBan
                             join k in _context.KipLamViec on nv.ID_Kip equals k.ID_Kip into ulist3
                             from k in ulist3.DefaultIfEmpty()
                             join vt in _context.ViTriLamViec on nv.ID_ViTri equals vt.ID_ViTri into ulist4
                             from vt in ulist4.DefaultIfEmpty()
                             join vtld in _context.ViTriLaoDong on a.ID_ViTriLaoDong equals vtld.ID_ViTriLaoDong into ulist5
                             from vtld in ulist5.DefaultIfEmpty()
                             select new KSK_BenhNgheNghiep
                             {
                                 ID_KSK_BNN = a.ID_KSK_BNN,
                                 ID_NV = (int)a.ID_NV,
                                 MaNV = nv.MaNV,
                                 HoTen = nv.HoTen,
                                 NgaySinh = (DateTime?)nv.NgaySinh ?? default,
                                 ID_PhongBan = (int?)a.ID_PhongBan??default,
                                 TenPhongBan = bp.TenPhongBan,
                                 TenKip = k.TenKip,
                                 TenViTri = vt.TenViTri,
                                 NgayKham = (DateTime?)a.NgayKham ?? default,
                                 NgayLenDanhSach = (DateTime?)a.NgayLenDanhSach ?? default,
                                 ID_ViTriLaoDong = (int)a.ID_ViTriLaoDong,
                                 TenViTriLaoDong = vtld.TenViTriLaoDong,
                                 GhiChu = a.GhiChu,
                                 ID_PheDuyet = (int?)a.ID_PheDuyet ?? default
                             }).ToListAsync();
            res = res.Where(x => x.NgayLenDanhSach == check.NgayTrinhKy).ToList();
            ViewBag.ID_TK = id;
            ViewBag.ID_PB = check.ID_PhongBan;
            var ct_nd = _context.CT_KSK_BenhNgheNghiep.ToList();
            ViewData["CT_KSK_BenhNgheNghiep"] = ct_nd;
            var ct_vt = _context.ViTriLaoDong.ToList();
            ViewData["ViTriLaoDong"] = ct_vt;
            var ct_tk = _context.TrinhKy.ToList();
            ViewData["TrinhKy"] = ct_tk;
            var ct_nv = _context.NhanVien.ToList();
            ViewData["NhanVien"] = ct_nv;
            var ct_bp = _context.PhongBan.ToList();
            ViewData["PhongBan"] = ct_bp;
            var ct_vtlv = _context.ViTriLamViec.ToList();
            ViewData["ViTriLamViec"] = ct_vtlv;
            var ct_ck = _context.TaiKhoan.ToList();
            ViewData["TaiKhoan"] = ct_ck;
            const int pageSize = 3000;
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
        private List<KSK_BenhNgheNghiep> GetDemarcation(int? ID_TK)
        {
            var check = _context.TrinhKy.Where(x => x.ID_TK == ID_TK).FirstOrDefault();
            var res = (from a in _context.KSK_BenhNgheNghiep.Where(x=>x.ID_PhongBan == check.ID_PhongBan && x.NgayLenDanhSach == check.NgayTrinhKy)
                            join nv in _context.NhanVien on a.ID_NV equals nv.ID_NV
                            join bp in _context.PhongBan on nv.ID_PhongBan equals bp.ID_PhongBan
                            join k in _context.KipLamViec on nv.ID_Kip equals k.ID_Kip into ulist3
                            from k in ulist3.DefaultIfEmpty()
                            join vt in _context.ViTriLamViec on nv.ID_ViTri equals vt.ID_ViTri into ulist4
                            from vt in ulist4.DefaultIfEmpty()
                            join vtld in _context.ViTriLaoDong on a.ID_ViTriLaoDong equals vtld.ID_ViTriLaoDong into ulist5
                            from vtld in ulist5.DefaultIfEmpty()
                            select new KSK_BenhNgheNghiep
                            {
                                ID_KSK_BNN = a.ID_KSK_BNN,
                                ID_NV = (int)a.ID_NV,
                                MaNV = nv.MaNV,
                                HoTen = nv.HoTen,
                                NgaySinh = (DateTime?)nv.NgaySinh ?? default,
                                NgayNhanViec =(DateTime?)nv.NgayVaoLam,
                                TenPhongBan = bp.TenPhongBan,
                                TenKip = k.TenKip,
                                TenViTri = vt.TenViTri,
                                NgayKham = (DateTime?)a.NgayKham ?? default,
                                NgayLenDanhSach = (DateTime?)a.NgayLenDanhSach ?? default,
                                ID_ViTriLaoDong = (int)a.ID_ViTriLaoDong,
                                TenViTriLaoDong = vtld.TenViTriLaoDong,
                                GhiChu = a.GhiChu,
                                ID_PheDuyet = (int?)a.ID_PheDuyet ?? default

                            }).ToList();
            return res;
        }

        public async Task<IActionResult> ExportToExcel(int? ID_TK)
        {
            try
            {
                var check = await _context.TrinhKy
                    .AsNoTracking()
                    .FirstOrDefaultAsync(x => x.ID_TK == ID_TK);

                if (check == null)
                    return BadRequest("Không tìm thấy nội dung trình ký");

                var res = await (
                    from a in _context.KSK_BenhNgheNghiep
                    where a.ID_PhongBan == check.ID_PhongBan
                    join nv in _context.NhanVien on a.ID_NV equals nv.ID_NV
                    join bp in _context.PhongBan on a.ID_PhongBan equals bp.ID_PhongBan
                    join k in _context.KipLamViec on nv.ID_Kip equals k.ID_Kip into k1
                    from k in k1.DefaultIfEmpty()
                    join vt in _context.ViTriLamViec on nv.ID_ViTri equals vt.ID_ViTri into vt1
                    from vt in vt1.DefaultIfEmpty()
                    join vtld in _context.ViTriLaoDong on a.ID_ViTriLaoDong equals vtld.ID_ViTriLaoDong into vtld1
                    from vtld in vtld1.DefaultIfEmpty()
                    select new
                    {
                        nv.MaNV,
                        nv.HoTen,
                        NgaySinh = nv.NgaySinh,
                        bp.TenPhongBan,
                        TenKip = k != null ? k.TenKip : "",
                        TenViTri = vt != null ? vt.TenViTri : "",
                        TenViTriLaoDong = vtld != null ? vtld.TenViTriLaoDong : "",
                        a.NgayKham,
                        a.NgayLenDanhSach,
                        a.GhiChu
                    }
                ).ToListAsync();

                res = res
                    .Where(x => x.NgayLenDanhSach == check.NgayTrinhKy)
                    .ToList();

                using var workbook = new XLWorkbook();
                var ws = workbook.Worksheets.Add("KSK_BNN");

                // ================= HEADER =================
                var headers = new[]
                {
                    "STT",
                    "Mã NV",
                    "Họ tên",
                    "Ngày sinh",
                    "Phòng ban",
                    "Kíp làm việc",
                    "Vị trí làm việc",
                    "Vị trí lao động",
                    "Ngày khám",
                    "Ghi chú"
                };

                for (int i = 0; i < headers.Length; i++)
                {
                    ws.Cell(1, i + 1).Value = headers[i];
                    ws.Cell(1, i + 1).Style.Font.Bold = true;
                    ws.Cell(1, i + 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    ws.Cell(1, i + 1).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                }

                // ================= DATA =================
                int row = 2;
                int stt = 1;

                foreach (var item in res)
                {
                    ws.Cell(row, 1).Value = stt++;
                    ws.Cell(row, 2).Value = item.MaNV;
                    ws.Cell(row, 3).Value = item.HoTen;
                    ws.Cell(row, 4).Value = item.NgaySinh;
                    ws.Cell(row, 5).Value = item.TenPhongBan;
                    ws.Cell(row, 6).Value = item.TenKip;
                    ws.Cell(row, 7).Value = item.TenViTri;
                    ws.Cell(row, 8).Value = item.TenViTriLaoDong;
                    ws.Cell(row, 9).Value = item.NgayKham;
                    ws.Cell(row, 10).Value = item.GhiChu;

                    ws.Cell(row, 4).Style.DateFormat.Format = "dd/MM/yyyy";
                    ws.Cell(row, 9).Style.DateFormat.Format = "dd/MM/yyyy";

                    row++;
                }

                ws.Columns().AdjustToContents();

                // ================= EXPORT =================
                using var stream = new MemoryStream();
                workbook.SaveAs(stream);

                return File(
                    stream.ToArray(),
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    $"DanhSach_KSK_BNN_{DateTime.Now:dd_MM_yyyy}.xlsx"
                );
            }
            catch (Exception ex)
            {
                TempData["msgError"] =
                    "<script>alert('Có lỗi khi xuất Excel. Vui lòng thử lại');</script>";

                return RedirectToAction("Index");
            }
        }
    }
}
