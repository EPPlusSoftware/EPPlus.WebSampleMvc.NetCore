using OfficeOpenXml;
using OfficeOpenXml.Export.HtmlExport;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace EPPlus.WebSampleMvc.NetCore.Models.HtmlExport
{
    public class ExportTable3Model
    {
        public void SetupSampleData(TableStyles style = TableStyles.Dark1)
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Html export sample 2");
                var csvFileInfo = new FileInfo(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"data\\currencies.csv"));
                var format = new ExcelTextFormat
                {
                    Delimiter = ';',
                    Culture = CultureInfo.InvariantCulture,
                    DataTypes = new eDataTypes[] { eDataTypes.DateTime, eDataTypes.Number, eDataTypes.Number, eDataTypes.Number }
                };
                var tableRange = sheet.Cells["A1"].LoadFromText(csvFileInfo, format, style, true);

                sheet.Cells[tableRange.Start.Row, 1, tableRange.End.Row, 1].Style.Numberformat.Format = "yyyy-MM-dd";
                sheet.Cells[tableRange.Start.Row, 3, tableRange.End.Row, 3].Style.Numberformat.Format = "#,##0.0000";
                sheet.Cells[tableRange.Start.Row, 4, tableRange.End.Row, 4].Style.Numberformat.Format = "#,##0.0000";

                var table = sheet.Tables.GetFromRange(tableRange);
                // export css and html
                Css = table.HtmlExporter.GetCssString();
                var o = HtmlTableExportOptions.Create();
                o.Culture = CultureInfo.InvariantCulture;
                Html = table.HtmlExporter.GetHtmlString(o);
            }
        }

        public string Css { get; set; }

        public string Html { get; set; }
    }
}
