﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using OfficeOpenXml;

namespace EPPlus.WebSampleMvc.NetCore.Controllers
{
    public class SamplesController : Controller
    {
        public SamplesController(IConfiguration configuration)
        {

        }

        public IActionResult Index()
        {
            return View();
        }

        /// <summary>
        /// Simple example on how to create a workbook without
        /// access to the file system and return it as a file.
        /// </summary>
        /// <returns>A workbook as a file with filename and content type set</returns>
        [HttpGet]
        public IActionResult GetWorkbook1()
        {
            using(var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet 1");
                sheet.Cells["A1:C1"].Merge = true;
                sheet.Cells["A1"].Style.Font.Size = 18f;
                sheet.Cells["A1"].Style.Font.Bold = true;
                sheet.Cells["A1"].Value = "Simple example 1";

                var excelData = package.GetAsByteArray();
                var contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                var fileName = "MyWorkbook.xlsx";
                return File(excelData, contentType, fileName);
            }
        }
    }
}