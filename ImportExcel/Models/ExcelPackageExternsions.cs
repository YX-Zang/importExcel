using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;


namespace ReadExcel.Models
{
    public static class ExcelPackageExtersions
    {

        public static DataTable ToDataTable(this ExcelPackage package)
        {
            ExcelWorksheet workSheet = package.Workbook.Worksheets.First();
            DataTable Dt = new DataTable();
            //抓成交年月
            //1.抓標題 2.抓銀行 --->到合計停止 so 思考for抓的範圍 要抓哪幾筆 (有值後 往右延伸N筆
            //散值該放哪裡? 可以放第二個table嗎
            foreach (var firstRowCell in workSheet.Cells[1,1,1,workSheet.Dimension.End.Column])
            {
                Dt.Columns.Add(firstRowCell.Text);
            }
            for (var rowNumber = 2; rowNumber <= workSheet.Dimension.End.Row;rowNumber++)
            {
                var row = workSheet.Cells[rowNumber, 1, rowNumber, workSheet.Dimension.End.Column];
                var newRow = Dt.NewRow();
                foreach(var cell in row)
                {
                    
                    newRow[cell.Start.Column - 1] = cell.Text;
                }
                Dt.Rows.Add(newRow);
            }
            return Dt;

        }
    }
}