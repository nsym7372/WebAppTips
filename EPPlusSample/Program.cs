using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusSample
{
    using System.IO;
    using OfficeOpenXml;
    class Program
    {
        static void Main(string[] args)
        {
            // https://chawatoyo.blog.fc2.com/blog-entry-15.html
            CellOperation();
            SheetOperation();
            BookOperation();
        }

        private static void BookOperation()
        {
            var sfi = new FileInfo("result.xlsx");
            var dfi = new FileInfo("result2.xlsx");
            if (dfi.Exists)
            {
                dfi.Delete();
                dfi = new FileInfo("result2.xlsx");
            }

            //シートを複製
            using (var src = new ExcelPackage(sfi))
            using (var dest = new ExcelPackage(dfi))
            {
                var target = src.Workbook.Names["CellName"].Worksheet;
                dest.Workbook.Worksheets.Add("sheet1");
                dest.Workbook.Worksheets.Add("シート名");
                dest.Workbook.Worksheets.Add("sheet2");

                //古いシートに設定された、名前付きセルのアドレスを退避
                //dest.Workbook.Names.Add("CellName", dest.Workbook.Worksheets["シート名"].Cells["A1"]);
                //var names = dest.Workbook.Names;

                //古いシート名を変更
                dest.Workbook.Worksheets[target.Name].Name = target.Name + "_";

                //新しいシートを挿入
                dest.Workbook.Worksheets.Add(target.Name, target);

                //新しいシートを古いシート横に移動
                dest.Workbook.Worksheets.MoveAfter(target.Name, target.Name + "_");

                //古いシートを削除
                dest.Workbook.Worksheets.Delete(target.Name + "_");

                

                dest.Save();
            }
        }

        private static void SheetOperation()
        {
            var fi = new FileInfo("result.xlsx");
            using (var pkg = new ExcelPackage(fi))
            {
                //名前付きのセルを取得、セルからシートを取得
                var sheet = pkg.Workbook.Names["CellName"].Worksheet;

                //B列（2列目）に3列追加
                sheet.InsertColumn(2, 3);

                pkg.Save();
            }
        }

        private static void CellOperation()
        {
            var newFile = new FileInfo("result.xlsx");
            if (newFile.Exists)
            {
                newFile.Delete();
                newFile = new FileInfo("result.xlsx");
            }

            using (ExcelPackage package = new ExcelPackage(newFile))
            {
                ExcelWorksheet sheet = package.Workbook.Worksheets.Add("シート名");
                sheet.Cells["A1"].Value = "Hello World";
                sheet.Cells[2, 1].Value = 27;
                sheet.Cells["A1"].Copy(sheet.Cells["A3"]);  //セルのコピー
                sheet.Cells["A1"].Offset(3, 0).Value = "offset"; //オフセット(A4)
                sheet.Cells["A5"].Value = sheet.Cells["A5"].Address; //アドレス
                sheet.Cells["A6"].Value = sheet.Cells["A6"].FullAddress;
                sheet.Cells["A7"].Value = sheet.Cells["A7"].Start.Row;  //行番号
                sheet.Cells["B7"].Value = sheet.Cells["B7"].Start.Column;   //列番号

                //セル結合して値入れる
                sheet.Cells["A8:B8"].Merge = true;
                sheet.Cells["A8"].Value = "Merged";

                //結合したセルの値を取得 https://stackoverflow.com/questions/47680167/get-merged-cell-area-with-epplus
                var thisCell = sheet.Cells["B8"];
                var i = thisCell.Worksheet.GetMergeCellId(thisCell.Start.Row, thisCell.Start.Column);
                var c = thisCell.Worksheet.MergedCells[i - 1];
                sheet.Cells["A9"].Value = thisCell.Worksheet.Cells[c].Value;

                //セルに名前付け
                package.Workbook.Names.Add("CellName", sheet.Cells["A1"]);

                //シートを複製、別名を付ける
                package.Workbook.Worksheets.Add("sheet2", sheet);

                //シートを新しく追加
                package.Workbook.Worksheets.Add("sheet3");

                package.Save();
            }
        }
    }
}
