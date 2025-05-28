using Microsoft.AspNetCore.Mvc;
using DataProfile.Models;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml;
using System.IO;
using Rotativa.AspNetCore;

namespace DataProfile.Controllers
{
    public class HomeController : Controller
    {
        private static List<Profile> _profiles = new List<Profile>()
        {
    new Profile { Id = 1, FullName = "Baso Arfan Efendy", Email = "basoarfane@gmail.com", PhoneNumber = "081234567890", BirthDate = new DateTime(1999, 5, 15), Address = "Makassar" },
    new Profile { Id = 2, FullName = "Jane Smith", Email = "jane.smith@example.com", PhoneNumber = "081234567891", BirthDate = new DateTime(1995, 8, 20), Address = "Jakarta" },
    new Profile { Id = 3, FullName = "Ahmad Haerul", Email = "ahmad.haerul@example.com", PhoneNumber = "081234567892", BirthDate = new DateTime(1992, 2, 10), Address = "Bandung" },
    new Profile { Id = 4, FullName = "Ilham Hidayat", Email = "ilham.hidayat@example.com", PhoneNumber = "081234567893", BirthDate = new DateTime(1998, 11, 5), Address = "Surabaya" },
    new Profile { Id = 5, FullName = "Siti Aminah", Email = "siti.aminah@example.com", PhoneNumber = "081234567894", BirthDate = new DateTime(1997, 7, 22), Address = "Medan" },
    new Profile { Id = 6, FullName = "Budi Santoso", Email = "budi.santoso@example.com", PhoneNumber = "081234567895", BirthDate = new DateTime(1990, 3, 12), Address = "Semarang" },
    new Profile { Id = 7, FullName = "Dewi Lestari", Email = "dewi.lestari@example.com", PhoneNumber = "081234567896", BirthDate = new DateTime(1996, 9, 30), Address = "Yogyakarta" },
    new Profile { Id = 8, FullName = "Rizky Pratama", Email = "rizky.pratama@example.com", PhoneNumber = "081234567897", BirthDate = new DateTime(2000, 1, 18), Address = "Palembang" },
    new Profile { Id = 9, FullName = "Fitriani Indah", Email = "fitriani.indah@example.com", PhoneNumber = "081234567898", BirthDate = new DateTime(1994, 4, 25), Address = "Denpasar" },
    new Profile { Id = 10, FullName = "Agus Setiawan", Email = "agus.setiawan@example.com", PhoneNumber = "081234567899", BirthDate = new DateTime(1991, 12, 3), Address = "Makassar" },
    new Profile { Id = 11, FullName = "Nurul Huda", Email = "nurul.huda@example.com", PhoneNumber = "081345678900", BirthDate = new DateTime(1993, 6, 8), Address = "Jakarta" },
    new Profile { Id = 12, FullName = "Eko Wijaya", Email = "eko.wijaya@example.com", PhoneNumber = "081345678901", BirthDate = new DateTime(1989, 10, 14), Address = "Bandung" },
    new Profile { Id = 13, FullName = "Putri Ayu", Email = "putri.ayu@example.com", PhoneNumber = "081345678902", BirthDate = new DateTime(1999, 8, 1), Address = "Surabaya" },
    new Profile { Id = 14, FullName = "Muhammad Zaki", Email = "muhammad.zaki@example.com", PhoneNumber = "081345678903", BirthDate = new DateTime(1997, 1, 28), Address = "Medan" },
    new Profile { Id = 15, FullName = "Lina Marlina", Email = "lina.marlina@example.com", PhoneNumber = "081345678904", BirthDate = new DateTime(1995, 5, 19), Address = "Semarang" },
    new Profile { Id = 16, FullName = "Dimas Anggara", Email = "dimas.anggara@example.com", PhoneNumber = "081345678905", BirthDate = new DateTime(1992, 11, 7), Address = "Yogyakarta" },
    new Profile { Id = 17, FullName = "Sri Wahyuni", Email = "sri.wahyuni@example.com", PhoneNumber = "081345678906", BirthDate = new DateTime(2001, 3, 3), Address = "Palembang" },
    new Profile { Id = 18, FullName = "Yoga Permana", Email = "yoga.permana@example.com", PhoneNumber = "081345678907", BirthDate = new DateTime(1996, 7, 11), Address = "Denpasar" },
    new Profile { Id = 19, FullName = "Cindy Claudia", Email = "cindy.claudia@example.com", PhoneNumber = "081345678908", BirthDate = new DateTime(1998, 9, 23), Address = "Makassar" },
    new Profile { Id = 20, FullName = "Fajar Nugroho", Email = "fajar.nugroho@example.com", PhoneNumber = "081345678909", BirthDate = new DateTime(1990, 12, 16), Address = "Jakarta" }

        };
        

        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        // GET: INDEX
        public IActionResult Index()
        {
            return View(_profiles.OrderBy(e => e.Id).ToList());
        }
        public IActionResult Details(int? id)
        {
            if (id == null)
            {
                return NotFound("ID tidak boleh kosong.");
            }
            var profile = _profiles.FirstOrDefault(m => m.Id == id);
            if (profile == null)
            {
                return NotFound($"Data dengan ID {id} tidak ditemukan.");
            }
            return View(profile); 
        }

        // Export to Excel
        public IActionResult ExportToExcel()
        {
            var profilesToExport = _profiles.ToList();
            if (!profilesToExport.Any())
            {
                TempData["ErrorMessage"] = "Tidak ada data profile untuk diekspor.";
                return RedirectToAction(nameof(Index));
            }

            OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;


            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Profiles");
                worksheet.Cells[1, 1].Value = "ID";
                worksheet.Cells[1, 2].Value = "Nama Lengkap";
                worksheet.Cells[1, 3].Value = "Email";
                worksheet.Cells[1, 4].Value = "No. Telepon";
                worksheet.Cells[1, 5].Value = "Tanggal Lahir";
                worksheet.Cells[1, 6].Value = "Alamat";

                using (var range = worksheet.Cells[1, 1, 1, 6])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
                }
                for (int i = 0; i < profilesToExport.Count; i++)
                {
                    var emp =  profilesToExport[i];
                    worksheet.Cells[i + 2, 1].Value = emp.Id;
                    worksheet.Cells[i + 2, 2].Value = emp.FullName;
                    worksheet.Cells[i + 2, 3].Value = emp.Email;
                    worksheet.Cells[i + 2, 4].Value = emp.PhoneNumber;
                    worksheet.Cells[i + 2, 5].Value = emp.BirthDate.ToString("dd/MM/yyyy");
                    worksheet.Cells[i + 2, 6].Value = emp.Address;
                }
                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                var stream = new MemoryStream();
                package.SaveAs(stream);
                stream.Position = 0;
                return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    $"Profile_Data_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
            }
        }

public IActionResult ExportToPdf()
{
    var profilesToExport = _profiles.ToList(); 
    if (!profilesToExport.Any())
    {
        TempData["ErrorMessage"] = "Tidak ada data profile untuk diekspor.";
        return RedirectToAction(nameof(Index));
    }

    // PERUBAHAN UTAMA dengan baris ini untuk memanggil Rotativa:
    return new ViewAsPdf("ProfilePdfView", profilesToExport)
    {
        FileName = $"Profile_Data_{DateTime.Now:yyyyMMdd_HHmmss}.pdf",
        PageSize = Rotativa.AspNetCore.Options.Size.A4,
        PageOrientation = Rotativa.AspNetCore.Options.Orientation.Landscape,
        // CustomSwitches = "--footer-center \"Halaman [page] dari [toPage]\" --footer-font-size \"8\""
    };
}
    }
}