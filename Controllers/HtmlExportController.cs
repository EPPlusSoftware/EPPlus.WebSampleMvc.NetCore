using EPPlus.WebSampleMvc.NetCore.Models.HtmlExport;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace EPPlus.WebSampleMvc.NetCore.Controllers
{
    public class HtmlExportController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }

        [HttpGet]
        public IActionResult ExportTable1()
        {
            var model = new ExportTable1Model();
            model.SetupSampleData(0);
            model.TableStyle = "Dark1";
            return View(model);
        }

        [HttpPost, ValidateAntiForgeryToken]
        public IActionResult ExportTable1(ExportTable1Model model)
        {
            if(!Enum.TryParse(model.TableStyle, out TableStyles ts))
            {
                ts = TableStyles.Dark1;
            }
            ViewData["TableStyle"] = ts.ToString();
            model.SetupSampleData(model.Theme, ts);
            return View(model);
        }

        public IActionResult ExportTable2(string style)
        {
            if (!Enum.TryParse(style, out TableStyles ts))
            {
                ts = TableStyles.Dark1;
            }
            ViewData["TableStyle"] = ts.ToString();
            var model = new ExportTable2Model();
            model.SetupSampleData(ts);
            return View(model);
        }

        public IActionResult ExportTable3(string style)
        {
            if (!Enum.TryParse(style, out TableStyles ts))
            {
                ts = TableStyles.Dark1;
            }
            var model = new ExportTable3Model();
            model.SetupSampleData(ts);
            return View(model);
        }
    }
}
