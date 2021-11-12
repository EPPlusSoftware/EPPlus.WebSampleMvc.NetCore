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

        public IActionResult ExportTable1(string style)
        {
            if(!Enum.TryParse(style, out TableStyles ts))
            {
                ts = TableStyles.Dark1;
            }
            ViewData["TableStyle"] = ts.ToString();
            var model = new ExportTable1Model();
            model.SetupSampleData(ts);
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
    }
}
