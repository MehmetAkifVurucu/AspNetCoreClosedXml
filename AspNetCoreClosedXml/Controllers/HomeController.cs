using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using AspNetCoreClosedXml.Models;
using ClosedXML.Excel;
using System.IO;

namespace AspNetCoreClosedXml.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }
        public IActionResult ExportExcelOgrenciList()
        {
            using (var workbook = new XLWorkbook())
            {
                var workSheet = workbook.Worksheets.Add("Öğrenci Listesi");

                workSheet.Cell(1, 1).Value = "Sınıf Adı";
                workSheet.Cell(1, 1).Style.Font.Bold = true;
                workSheet.Cell(1, 2).Value = "12/A";

                workSheet.Cell(2, 1).Value = "Oluşturma Tarihi";
                workSheet.Cell(2, 1).Style.Font.Bold = true;
                workSheet.Cell(2, 2).SetValue<string>(Convert.ToString(DateTime.Now));

                workSheet.Cell(3, 1).Value = "Öğrenci No";
                workSheet.Cell(3, 2).Value = "Ad Soyad";

                for (int i = 1; i <= 2; i++)
                {
                    workSheet.Cell(3, i).Style.Font.Bold = true;
                }

                int ogrenciSayac = 4;
                foreach (var item in GetOgrenciList())
                {
                    workSheet.Cell(ogrenciSayac, 1).Value = item.OgrenciNo;
                    workSheet.Cell(ogrenciSayac, 2).Value = item.AdSoyad;
                    ogrenciSayac++;
                }

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();
                    return File(
                                content,
                                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                "12A Sınıfı Öğrenci Listesi.xlsx"
                                );
                }
            }
        }


        private List<OgrenciModel> GetOgrenciList()
        {
            List<OgrenciModel> ogrencis = new List<OgrenciModel>
            {
             new OgrenciModel { OgrenciNo = 1, AdSoyad = "Ali Mutlu"},
             new OgrenciModel { OgrenciNo = 2, AdSoyad = "Ceren Kaya"},
             new OgrenciModel { OgrenciNo = 3, AdSoyad = "Mehmet Çekirdek"},
             new OgrenciModel { OgrenciNo = 4, AdSoyad = "Hakan Taşçı"},
             new OgrenciModel { OgrenciNo = 5, AdSoyad = "Gizim Gülenyüz" },
             new OgrenciModel { OgrenciNo = 6, AdSoyad = "Hasan Özgür" }

            };
            return ogrencis;
        }
    }
}
