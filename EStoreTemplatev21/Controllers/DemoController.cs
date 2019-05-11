using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DinkToPdf;
using DinkToPdf.Contracts;
using EStoreTemplatev21.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;

namespace EStoreTemplatev21.Controllers
{
    public class DemoController : Controller
    {
        private readonly MyeStoreContext ctx;
        private IConverter _converter;
        public DemoController(MyeStoreContext db, IConverter con)
        {
            ctx = db; _converter = con;
        }

        public IActionResult CreatePDF()
        {
            var globalSettings = new GlobalSettings
            {
                ColorMode = ColorMode.Color,
                Orientation = Orientation.Portrait,
                PaperSize = PaperKind.A4,
                Margins = new MarginSettings { Top = 10 },
                DocumentTitle = "Loai Report"
            };

            var objectSettings = new ObjectSettings
            {
                PagesCount = true,
                HtmlContent = TemplateGenerator.GetHTMLString(),
                WebSettings = { DefaultEncoding = "utf-8", UserStyleSheet = Path.Combine(Directory.GetCurrentDirectory(), "assets", "styles.css") },
                HeaderSettings = { FontName = "Arial", FontSize = 9, Right = "Trang [page]/[toPage]", Line = true },
                FooterSettings = { FontName = "Arial", FontSize = 9, Line = true, Center = "Report Footer" }
            };

            var pdf = new HtmlToPdfDocument()
            {
                GlobalSettings = globalSettings,
                Objects = { objectSettings }
            };

            var file = _converter.Convert(pdf);
            return File(file, "application/pdf");
        }

        public IActionResult Export()
        {
            //chuẩn bị dữ liệu xuất
            var data = ctx.Loai.ToList();

            var stream = new MemoryStream();

            using (var package = new ExcelPackage(stream))
            {
                var sheet = package.Workbook.Worksheets.Add("Loai");
                sheet.Cells.LoadFromCollection(data, true);
                //sheet.Cells[1, 1].Value = "Mã loại";
                package.Save();
            }

            stream.Position = 0;
            string fileName = $"Loai_{DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss")}.xlsx";

            return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }

        public IActionResult ImportLoai()
        {
            return View();
        }

        [HttpPost]
        public IActionResult ImportLoai(IFormFile fImport)
        {
            if (fImport == null || fImport.Length <= 0)
            {
                ViewBag.ThongBao = "File không tồn tại hoặc có lỗi upload";
                return View();
            }

            List<Loai> loaiImports = new List<Loai>();

            //tạo stream giữ file upload lên
            using (var stream = new MemoryStream())
            {
                fImport.CopyTo(stream);

                //Map stream với Excel file
                using (var package = new ExcelPackage(stream))
                {
                    var sheet = package.Workbook.Worksheets[0];
                    int rowCount = sheet.Dimension.Rows;

                    //duyệt qua từng dòng của sheet Excel bóc tách dữ liệu ra
                    for(int i = 2; i <= rowCount; i++)
                    {
                        loaiImports.Add(new Loai {
                            MaLoai = int.Parse(sheet.Cells[i, 1].Value.ToString()),
                            TenLoai = sheet.Cells[i, 2].Value.ToString(),
                            MoTa = sheet.Cells[i, 3].Value.ToString(),
                            Hinh = sheet.Cells[i, 4].Value.ToString()                            
                        });
                    }
                }
            }

            if(loaiImports.Count > 0)
            {
                ctx.Database.OpenConnection();
                //tắt tự tăng
                ctx.Database.ExecuteSqlCommand("SET IDENTITY_INSERT dbo.Loai ON");
                ctx.SaveChanges();
                //tiến hành update hoặc insert
                foreach (Loai lo in loaiImports)
                {
                    var item = ctx.Loai.SingleOrDefault(p => p.MaLoai == lo.MaLoai);
                    if(item != null)//đã có --> update
                    {
                        item.TenLoai = lo.TenLoai;
                        item.MoTa = lo.MoTa;
                        item.Hinh = lo.Hinh;
                    }
                    else
                    {
                        ctx.Add(lo);
                    }
                }
                ctx.SaveChanges();
                //bật tự tăng
                ctx.Database.ExecuteSqlCommand("SET IDENTITY_INSERT dbo.Loai OFF");
                ctx.SaveChanges();
                ctx.Database.CloseConnection();
            }
            ViewBag.ThongBao = "Import thành công";
            return View();
        }

        public IActionResult Index()
        {
            return View();
        }
    }
}