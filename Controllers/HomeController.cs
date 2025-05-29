
using Microsoft.AspNetCore.Mvc;
using DataProfile.Models;         
using DataProfile.DataAccessLayer; 
using OfficeOpenXml;
using Rotativa.AspNetCore;
using System.IO;
using System.Linq;
using System.Collections.Generic;

namespace DataProfile.Controllers 
{
    public class HomeController : Controller
    {
        private readonly IProfileRepository _profileRepository;
        private readonly ILogger<HomeController> _logger;

        public HomeController(IProfileRepository profileRepository, ILogger<HomeController> logger)
        {
            _profileRepository = profileRepository;
            _logger = logger;
        }

        public IActionResult Index()
        {
            try
            {
                // Mengambil data dari repository (yang akan memanggil Stored Procedure)
                var profiles = _profileRepository.GetAllProfiles();
                return View(profiles); 
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error occurred while retrieving profiles for Index.");
                ViewBag.ErrorMessage = "Terjadi kesalahan saat mengambil data profil dari database.";
                return View(new List<Profile>()); 
            }
        }

        public IActionResult Details(int? id)
        {
            if (id == null)
            {
                _logger.LogWarning("Details action called with null ID.");
                return NotFound("ID tidak boleh kosong.");
            }
            try
            {
                // Mengambil data dari repository
                var profiles = _profileRepository.GetAllProfiles();
                var profile = profiles.FirstOrDefault(p => p.Id == id);


                if (profile == null)
                {
                    _logger.LogWarning($"Profile with ID {id} not found via repository.");
                    return NotFound($"Data dengan ID {id} tidak ditemukan.");
                }
                return View(profile); 
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error retrieving profile with ID {id}.");
                ViewBag.ErrorMessage = "Terjadi kesalahan saat mengambil detail profil.";
                return RedirectToAction(nameof(Index)); 
            }
        }


        public IActionResult ExportToExcel()
        {
            List<Profile> profilesToExport;
            try
            {
                profilesToExport = _profileRepository.GetAllProfiles();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error retrieving profiles for Excel export.");
                TempData["ErrorMessage"] = "Gagal mengambil data untuk ekspor Excel.";
                return RedirectToAction(nameof(Index));
            }
            
            if (profilesToExport == null || !profilesToExport.Any())
            {
                TempData["ErrorMessage"] = "Tidak ada data profil untuk diekspor.";
                return RedirectToAction(nameof(Index));
            }

            OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Profiles");
                // Headers
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
                    var prof = profilesToExport[i];
                    worksheet.Cells[i + 2, 1].Value = prof.Id;
                    worksheet.Cells[i + 2, 2].Value = prof.FullName;
                    worksheet.Cells[i + 2, 3].Value = prof.Email;
                    worksheet.Cells[i + 2, 4].Value = prof.PhoneNumber;
                    worksheet.Cells[i + 2, 5].Value = prof.BirthDate.ToString("dd/MM/yyyy");
                    worksheet.Cells[i + 2, 6].Value = prof.Address;
                }
                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                var stream = new MemoryStream();
                package.SaveAs(stream);
                stream.Position = 0;
                return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    $"Profil_Data_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
            }
        }

        public IActionResult ExportToPdf()
        {
            List<Profile> profilesToExport;
            try
            {
                profilesToExport = _profileRepository.GetAllProfiles();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error retrieving profiles for PDF export.");
                TempData["ErrorMessage"] = "Gagal mengambil data untuk ekspor PDF.";
                return RedirectToAction(nameof(Index));
            }

            if (profilesToExport == null || !profilesToExport.Any())
            {
                TempData["ErrorMessage"] = "Tidak ada data profil untuk diekspor.";
                return RedirectToAction(nameof(Index));
            }

          
            return new ViewAsPdf("ProfilePdfView", profilesToExport) 
            {
                FileName = $"Profil_Data_{DateTime.Now:yyyyMMdd_HHmmss}.pdf",
                PageOrientation = Rotativa.AspNetCore.Options.Orientation.Landscape
            };
        }
    }
}