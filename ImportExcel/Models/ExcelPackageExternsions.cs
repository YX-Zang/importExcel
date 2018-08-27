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

            //*****處理好標題遺漏的問題******
            workSheet.Cells[9, 15].Value = "占公司成交比率(筆數)";
            workSheet.Cells[9, 16].Value = "占公司成交比率(金額)";
            workSheet.Cells[9, 17].Value = "占市場成交比率(筆數)";
            workSheet.Cells[9, 18].Value = "占市場成交比率(金額)";
            foreach (var firstRowCell in workSheet.Cells[9, 1, 1, workSheet.Dimension.End.Column])
            {
                Dt.Columns.Add(firstRowCell.Text);
            }
            
            // 從11開始 一直向下探 
            // workSheet.Dimension.End.Row -2 把最後兩行備註忽略不看
            for (var rowNumber = 11; rowNumber <= workSheet.Dimension.End.Row -2; rowNumber++)
            {
                //這邊應該要修改為 向下一直判斷 有null跳過 沒有null向右延伸19次
                //抓出目前的Excel列
                ExcelRange range = workSheet.Cells[rowNumber, 1];
                if (range.Any(c => !string.IsNullOrEmpty(c.Text)) == true)
                    //這是一個完全空白列(使用者用Delete鍵刪除動作)
                    //空白有兩種 一種是使用者delete的空白，另一種是起初就沒有值
                {
                    //從rowNumber = 11 開始 一直向右探
                    var row = workSheet.Cells[rowNumber, 1, rowNumber, workSheet.Dimension.End.Column];
                    var newRow = Dt.NewRow();

                    foreach (var cell in row)
                    {
                        if (cell.Start.Column == 1)
                        {
                            newRow[cell.Start.Column - 1] = cell.Style.Numberformat.Format = "0";
                        }
                        else if (cell.Start.Column == 4 || cell.Start.Column == 5 || cell.Start.Column == 6 || cell.Start.Column == 8 || cell.Start.Column == 9 || cell.Start.Column == 10 || cell.Start.Column == 11 || cell.Start.Column == 12 || cell.Start.Column == 13 || cell.Start.Column == 14)
                        {
                            newRow[cell.Start.Column - 1] = cell.Style.Numberformat.Format = "#,##0";
                        }
                        else if (cell.Start.Column == 15 || cell.Start.Column == 16 || cell.Start.Column == 17 || cell.Start.Column == 18 || cell.Start.Column == 19)
                        {
                            newRow[cell.Start.Column - 1] = cell.Style.Numberformat.Format = "0.00%";
                        }
                        newRow[cell.Start.Column - 1] = cell.Text;
                        //newRow[cell.Start.Column - 1] = cell.RichText != null ? cell.RichText.Text : cell.Text;
                    }

                    Dt.Rows.Add(newRow);

                }
            }
            return Dt;

        }
    }
}