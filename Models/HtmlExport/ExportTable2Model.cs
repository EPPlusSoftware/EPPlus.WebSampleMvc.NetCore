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
    public class ExportTable2Model
    {
        public ExportTable2Model()
        {

        }

        private DataTable _dataTable;

        private void InitDataTable()
        {
            _dataTable = new DataTable();
            _dataTable.Columns.Add("Country", typeof(string));
            _dataTable.Columns.Add("Population", typeof(int));


            _dataTable.Rows.Add("Sweden", 10409248);
            _dataTable.Rows.Add("Norway", 5402171);
            _dataTable.Rows.Add("Netherlands", 17553530);
            _dataTable.Rows.Add("Finland", 5541806);
            _dataTable.Rows.Add("Belgium", 11521238);
            _dataTable.Rows.Add("Denmark", 5850189);
            _dataTable.Rows.Add("Lithuania", 2801264);
            _dataTable.Rows.Add("Greece", 10718565);
            _dataTable.Rows.Add("Russia", 145734038);
            _dataTable.Rows.Add("Germany", 83124418);
            _dataTable.Rows.Add("France", 64990511);
            _dataTable.Rows.Add("Czech Republic", 10665677);
            _dataTable.Rows.Add("Slovakia", 5459781);
            _dataTable.Rows.Add("Spain", 47394223);
            _dataTable.Rows.Add("Portugal", 10256193);
            _dataTable.Rows.Add("United Kingdom", 67141684);
            _dataTable.Rows.Add("Poland", 37921592);
            _dataTable.Rows.Add("Albania", 2882740);
            _dataTable.Rows.Add("Estonia", 1322920);
            _dataTable.Rows.Add("Hungary", 9707499);
            _dataTable.Rows.Add("Romania", 19186000);
            _dataTable.Rows.Add("Italy", 60627291);
            _dataTable.Rows.Add("Bulgaria", 7051608);
            _dataTable.Rows.Add("Belarus", 9452617);
            _dataTable.Rows.Add("Austria", 8891388);
            _dataTable.Rows.Add("Switzerland", 8525611);
            _dataTable.Rows.Add("Ireland", 4818690);
            _dataTable.Rows.Add("Ukraine", 44246156);
            _dataTable.Rows.Add("Iceland", 336713);
            _dataTable.Rows.Add("Serbia", 6871547);
            _dataTable.Rows.Add("Croatia", 4156405);
            _dataTable.Rows.Add("Latvia", 1928459);
            _dataTable.Rows.Add("Bosnia and Herzegovina", 3323925);
            _dataTable.Rows.Add("Montenegro", 627809);
            _dataTable.Rows.Add("Cyrprus", 1189265);
            _dataTable.Rows.Add("Kosovo", 1798506);
            _dataTable.Rows.Add("Georgia", 1798506);

        }

        public void SetupSampleData(TableStyles style = TableStyles.Dark1)
        {
            InitDataTable();
            using(var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Html export sample 2");
                var tableRange = sheet.Cells["A1"].LoadFromDataTable(_dataTable, true, style);
                
                // sort the table range
                tableRange.Sort(x => x.SortBy.Column(1, eSortOrder.Descending));
                
                // configure the table
                var table = sheet.Tables.GetFromRange(tableRange);
                table.ShowTotal = true;
                table.Columns[1].TotalsRowFunction = RowFunctions.Sum;
                sheet.Calculate();

                // export css and html
                Css = table.HtmlExporter.GetCssString();
                Html = table.HtmlExporter.GetHtmlString();
            }
        }

        public string Css { get; set; }

        public string Html { get; set; }


    }
}
