using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using Faker;
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

        public void SetupSampleData(TableStyles style = TableStyles.Dark1)
        {
            InitDataTable();
            using(var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Html export sample 1");
                var tableRange = sheet.Cells["A1"].LoadFromDataTable(_dataTable, true, style);
                sheet.Cells["C2:C52"].Style.Numberformat.Format = "yyyy-MM-dd";
                var table = sheet.Tables.GetFromRange(tableRange);
                Css = table.HtmlExporter.GetCssString();
                Html = table.HtmlExporter.GetHtmlString();
            }
        }

        public string Css { get; set; }

        public string Html { get; set; }


    }
}
