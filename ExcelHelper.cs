using BookStates.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace BookStates.Helpers
{
    public class ExcelHelper
    {
       static public void Export(List<BookFromWeb> LitBooks, List<BookFromWeb> SabaBooks, List<BookFromWebForCommon> CommonBooks)
        {

            ExcelPackage excel = new ExcelPackage();
            var LitworkSheet = excel.Workbook.Worksheets.Add("lit.ge");
            var SabaworkSheet = excel.Workbook.Worksheets.Add("saba.com.ge");
            var CommonworkSheet = excel.Workbook.Worksheets.Add("common books");
            LitworkSheet.Cells[1, 1].LoadFromCollection(LitBooks, true);
            SabaworkSheet.Cells[1, 1].LoadFromCollection(SabaBooks, true);
            CommonworkSheet.Cells[1, 1].LoadFromCollection(CommonBooks, true);
            using (var memoryStream = new MemoryStream())
            {
                HttpContext.Current.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                HttpContext.Current.Response.AddHeader("content-disposition", "attachment;  filename=Books_Data.xlsx");
                excel.SaveAs(memoryStream);
                memoryStream.WriteTo(HttpContext.Current.Response.OutputStream);
                HttpContext.Current.Response.Flush();
                HttpContext.Current.Response.End();
            }
        }
    }
}