using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Faker;
using Microsoft.AspNetCore.Mvc.Rendering;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace EPPlus.WebSampleMvc.NetCore.Models.HtmlExport
{
    public class ExportTable1Model
    {
        public ExportTable1Model()
        {

        }

        private DataTable _dataTable;

        private void InitDataTable()
        {
            _dataTable = new DataTable();
            _dataTable.Columns.Add("Country", typeof(string));
            _dataTable.Columns.Add("FirstName", typeof(string));
            _dataTable.Columns.Add("LastName", typeof(string));
            _dataTable.Columns.Add("BirthDate", typeof(DateTime));
            _dataTable.Columns.Add("City", typeof(string));
            

            for(var x = 0; x < 50; x++)
            {
                _dataTable.Rows.Add(Faker.Address.UkCountry(), Faker.Name.First(), Faker.Name.Last(), Faker.Identification.DateOfBirth(), Faker.Address.City());
            }

        }

        public void SetupSampleData(int theme, TableStyles? style = TableStyles.Dark1)
        {
            InitDataTable();
            using(var package = new ExcelPackage())
            {
                if(theme > 0)
                {
                    var fileName = string.Empty;
                    switch(theme)
                    {
                        case 1:
                            fileName = "Ion";
                            break;
                        case 2:
                            fileName = "Banded";
                            break;
                        case 3:
                            fileName = "Parallax";
                            break;
                        default:
                            fileName = "Ion";
                            break;
                    }
                    var fi = new FileInfo(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"themes\\{fileName}.thmx"));
                    package.Workbook.ThemeManager.Load(fi);
                }
                
                var sheet = package.Workbook.Worksheets.Add("Html export sample 1");
                var tableRange = sheet.Cells["A1"].LoadFromDataTable(_dataTable, true, style);
                sheet.Cells["D2:D52"].Style.Numberformat.Format = "yyyy-MM-dd";
                tableRange.AutoFitColumns();
                
                var table = sheet.Tables.GetFromRange(tableRange);
                
                // table properties
                table.ShowFirstColumn = ShowFirstColumn;
                table.ShowLastColumn = ShowLastColumn;
                table.ShowColumnStripes = ShowColumnStripes;
                table.ShowRowStripes = ShowRowsStripes;

                Css = table.HtmlExporter.GetCssString();
                Html = table.HtmlExporter.GetHtmlString();
                WorkbookBytes = package.GetAsByteArray();
            }
        }

        public IEnumerable<SelectListItem> AllBuiltInTableStyles
        {
            get
            {
                return System.Enum.GetValues(typeof(TableStyles))
                    .Cast<TableStyles>()
                    .Where(x => x != TableStyles.Custom)
                    .Select(x => new SelectListItem(x.ToString(), x.ToString()));
            }
        }

        public IEnumerable<SelectListItem> AllThemes
        {
            get
            {
                return new List<SelectListItem>
                {
                    new SelectListItem("Default (Office)", "0"),
                    new SelectListItem("Ion", "1"),
                    new SelectListItem("Banded", "2"),
                    new SelectListItem("Parallax", "3")
                };
            }
        }

        public bool ShowFirstColumn { get; set; }

        public bool ShowLastColumn { get; set; }

        public bool ShowColumnStripes { get; set; }

        public bool ShowRowsStripes { get; set; }

        public string TableStyle { get; set; }

        public int Theme { get; set; }

        public bool AddBootstrapClasses { get; set; }

        public bool AddDataTablesJs { get; set; }

        public bool GetWorkbook { get; set; }

        public string Css { get; set; }

        public string Html { get; set; }

        public byte[] WorkbookBytes { get; set; }


    }
}
