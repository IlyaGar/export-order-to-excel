using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using WebAppExcelReport.Models;

namespace WebAppExcelReport.Services
{
    public class WriterExcel : IWriterExcel
    {
        private FullOrder _fullOrder = new FullOrder
        {
            order = new Order { id = 1232465, creationDate = "21.01.2021", deparmentId = "Инструменты", statusId = "Черновик", storeId = "Долгиновский", author = "Иванов" },
            orderBodies = new List<OrderBody>
            { 
                new OrderBody { group = "2.1.3.", article = "12345678", barcode = "100000121", name = "HJKKHKAflkamn HJKKHKAflkamn HJKKHKAflkamn HJKKHKAflkamn HJKKHKAflkamn v HJKKHKAflkamn", supplier = "OAO JJFJF", 
                    goods = "1", average = "4", stockDay = "5", deliveryDate = "21.01.2021", managerСomment = "skvjklsj", departmentComment = "svfkdsj" },
                new OrderBody { group = "2.1.2-", article = "12345678", barcode = "100000121", name = "HJKKHKAflkamn", supplier = "OAO JJFJF",
                    goods = "1", average = "4", stockDay = "5", deliveryDate = "21.01.2021", managerСomment = "skvjklsj", departmentComment = "svfkdsj" },
                new OrderBody { group = "2.1.3-", article = "12345678", barcode = "100000121", name = "HJKKHKAflkamn", supplier = "OAO JJFJF",
                    goods = "1", average = "4", stockDay = "5", deliveryDate = "21.01.2021", managerСomment = "skvjklsj", departmentComment = "svfkdsj" },
                new OrderBody { group = "2.1.3.", article = "12345678", barcode = "100000121", name = "HJKKHKAflkamn", supplier = "OAO JJFJF",
                    goods = "1", average = "4", stockDay = "5", deliveryDate = "21.01.2021", managerСomment = "skvjklsj", departmentComment = "svfkdsj" },
                new OrderBody { group = "2.1.2-", article = "12345678", barcode = "100000121", name = "HJKKHKAflkamn", supplier = "OAO JJFJF",
                    goods = "1", average = "4", stockDay = "5", deliveryDate = "21.01.2021", managerСomment = "skvjklsj", departmentComment = "svfkdsj" },
                new OrderBody { group = "2.1.3-", article = "12345678", barcode = "100000121", name = "HJKKHKAflkamn", supplier = "OAO JJFJF",
                    goods = "1", average = "4", stockDay = "5", deliveryDate = "21.01.2021", managerСomment = "skvjklsj", departmentComment = "svfkdsj" },
                new OrderBody { group = "2.1.3.", article = "12345678", barcode = "100000121", name = "HJKKHKAflkamn", supplier = "OAO JJFJF",
                    goods = "1", average = "4", stockDay = "5", deliveryDate = "21.01.2021", managerСomment = "skvjklsj", departmentComment = "svfkdsj" },
                new OrderBody { group = "2.1.2-", article = "12345678", barcode = "100000121", name = "HJKKHKAflkamn", supplier = "OAO JJFJF",
                    goods = "1", average = "4", stockDay = "5", deliveryDate = "21.01.2021", managerСomment = "skvjklsj", departmentComment = "svfkdsj" },
                new OrderBody { group = "2.1.3-", article = "12345678", barcode = "100000121", name = "HJKKHKAflkamn", supplier = "OAO JJFJF",
                    goods = "1", average = "4", stockDay = "5", deliveryDate = "21.01.2021", managerСomment = "skvjklsj", departmentComment = "svfkdsj" },
                new OrderBody { group = "2.1.3.", article = "12345678", barcode = "100000121", name = "HJKKHKAflkamn", supplier = "OAO JJFJF",
                    goods = "1", average = "4", stockDay = "5", deliveryDate = "21.01.2021", managerСomment = "skvjklsj", departmentComment = "svfkdsj" },
                new OrderBody { group = "2.1.2-", article = "12345678", barcode = "100000121", name = "HJKKHKAflkamn", supplier = "OAO JJFJF",
                    goods = "1", average = "4", stockDay = "5", deliveryDate = "21.01.2021", managerСomment = "skvjklsj", departmentComment = "svfkdsj" },
                new OrderBody { group = "2.1.3-", article = "12345678", barcode = "100000121", name = "HJKKHKAflkamn", supplier = "OAO JJFJF",
                    goods = "1", average = "4", stockDay = "5", deliveryDate = "21.01.2021", managerСomment = "skvjklsj", departmentComment = "svfkdsj" },
                new OrderBody { group = "2.1.3.", article = "12345678", barcode = "100000121", name = "HJKKHKAflkamn", supplier = "OAO JJFJF",
                    goods = "1", average = "4", stockDay = "5", deliveryDate = "21.01.2021", managerСomment = "skvjklsj", departmentComment = "svfkdsj" },
                new OrderBody { group = "2.1.2-", article = "12345678", barcode = "100000121", name = "HJKKHKAflkamn", supplier = "OAO JJFJF",
                    goods = "1", average = "4", stockDay = "5", deliveryDate = "21.01.2021", managerСomment = "skvjklsj", departmentComment = "svfkdsj" },
                new OrderBody { group = "2.1.3-", article = "12345678", barcode = "100000121", name = "HJKKHKAflkamn", supplier = "OAO JJFJF",
                    goods = "1", average = "4", stockDay = "5", deliveryDate = "21.01.2021", managerСomment = "skvjklsj", departmentComment = "svfkdsj" },
                new OrderBody { group = "2.1.3.", article = "12345678", barcode = "100000121", name = "HJKKHKAflkamn", supplier = "OAO JJFJF",
                    goods = "1", average = "4", stockDay = "5", deliveryDate = "21.01.2021", managerСomment = "skvjklsj", departmentComment = "svfkdsj" },
                new OrderBody { group = "2.1.2-", article = "12345678", barcode = "100000121", name = "HJKKHKAflkamn", supplier = "OAO JJFJF",
                    goods = "1", average = "4", stockDay = "5", deliveryDate = "21.01.2021", managerСomment = "skvjklsj", departmentComment = "svfkdsj" },
                new OrderBody { group = "2.1.3-", article = "12345678", barcode = "100000121", name = "HJKKHKAflkamn", supplier = "OAO JJFJF",
                    goods = "1", average = "4", stockDay = "5", deliveryDate = "21.01.2021", managerСomment = "skvjklsj", departmentComment = "svfkdsj" },
                new OrderBody { group = "2.1.3.", article = "12345678", barcode = "100000121", name = "HJKKHKAflkamn", supplier = "OAO JJFJF",
                    goods = "1", average = "4", stockDay = "5", deliveryDate = "21.01.2021", managerСomment = "skvjklsj", departmentComment = "svfkdsj" },
                new OrderBody { group = "2.1.2-", article = "12345678", barcode = "100000121", name = "HJKKHKAflkamn", supplier = "OAO JJFJF",
                    goods = "1", average = "4", stockDay = "5", deliveryDate = "21.01.2021", managerСomment = "skvjklsj", departmentComment = "svfkdsj" },
                new OrderBody { group = "2.1.3-", article = "12345678", barcode = "100000121", name = "HJKKHKAflkamn", supplier = "OAO JJFJF",
                    goods = "1", average = "4", stockDay = "5", deliveryDate = "21.01.2021", managerСomment = "skvjklsj", departmentComment = "svfkdsj" },
                new OrderBody { group = "2.1.3.", article = "12345678", barcode = "100000121", name = "HJKKHKAflkamn", supplier = "OAO JJFJF",
                    goods = "1", average = "4", stockDay = "5", deliveryDate = "21.01.2021", managerСomment = "skvjklsj", departmentComment = "svfkdsj" },
                new OrderBody { group = "2.1.2-", article = "12345678", barcode = "100000121", name = "HJKKHKAflkamn", supplier = "OAO JJFJF",
                    goods = "1", average = "4", stockDay = "5", deliveryDate = "21.01.2021", managerСomment = "skvjklsj", departmentComment = "svfkdsj" },
                new OrderBody { group = "2.1.3-", article = "12345678", barcode = "100000121", name = "HJKKHKAflkamn", supplier = "OAO JJFJF",
                    goods = "1", average = "4", stockDay = "5", deliveryDate = "21.01.2021", managerСomment = "skvjklsj", departmentComment = "svfkdsj" },
                new OrderBody { group = "2.1.3.", article = "12345678", barcode = "100000121", name = "HJKKHKAflkamn", supplier = "OAO JJFJF",
                    goods = "1", average = "4", stockDay = "5", deliveryDate = "21.01.2021", managerСomment = "skvjklsj", departmentComment = "svfkdsj" },
                new OrderBody { group = "2.1.2-", article = "12345678", barcode = "100000121", name = "HJKKHKAflkamn", supplier = "OAO JJFJF",
                    goods = "1", average = "4", stockDay = "5", deliveryDate = "21.01.2021", managerСomment = "skvjklsj", departmentComment = "svfkdsj" },
                new OrderBody { group = "2.1.3-", article = "12345678", barcode = "100000121", name = "HJKKHKAflkamn", supplier = "OAO JJFJF",
                    goods = "1", average = "4", stockDay = "5", deliveryDate = "21.01.2021", managerСomment = "skvjklsj", departmentComment = "svfkdsj" },
                new OrderBody { group = "2.1.3.", article = "12345678", barcode = "100000121", name = "HJKKHKAflkamn", supplier = "OAO JJFJF",
                    goods = "1", average = "4", stockDay = "5", deliveryDate = "21.01.2021", managerСomment = "skvjklsj", departmentComment = "svfkdsj" },
                new OrderBody { group = "2.1.2-", article = "12345678", barcode = "100000121", name = "HJKKHKAflkamn", supplier = "OAO JJFJF",
                    goods = "1", average = "4", stockDay = "5", deliveryDate = "21.01.2021", managerСomment = "skvjklsj", departmentComment = "svfkdsj" },
                new OrderBody { group = "2.1.3-", article = "12345678", barcode = "100000121", name = "HJKKHKAflkamn", supplier = "OAO JJFJF",
                    goods = "1", average = "4", stockDay = "5", deliveryDate = "21.01.2021", managerСomment = "skvjklsj", departmentComment = "svfkdsj" },
                new OrderBody { group = "2.1.3.", article = "12345678", barcode = "100000121", name = "HJKKHKAflkamn", supplier = "OAO JJFJF",
                    goods = "1", average = "4", stockDay = "5", deliveryDate = "21.01.2021", managerСomment = "skvjklsj", departmentComment = "svfkdsj" },
                new OrderBody { group = "2.1.2-", article = "12345678", barcode = "100000121", name = "HJKKHKAflkamn", supplier = "OAO JJFJF",
                    goods = "1", average = "4", stockDay = "5", deliveryDate = "21.01.2021", managerСomment = "skvjklsj", departmentComment = "svfkdsj" },
                new OrderBody { group = "2.1.3-", article = "12345678", barcode = "100000121", name = "HJKKHKAflkamn", supplier = "OAO JJFJF",
                    goods = "1", average = "4", stockDay = "5", deliveryDate = "21.01.2021", managerСomment = "skvjklsj", departmentComment = "svfkdsj" },
                new OrderBody { group = "2.1.3.", article = "12345678", barcode = "100000121", name = "HJKKHKAflkamn", supplier = "OAO JJFJF",
                    goods = "1", average = "4", stockDay = "5", deliveryDate = "21.01.2021", managerСomment = "skvjklsj", departmentComment = "svfkdsj" },
                new OrderBody { group = "2.1.2-", article = "12345678", barcode = "100000121", name = "HJKKHKAflkamn", supplier = "OAO JJFJF",
                    goods = "1", average = "4", stockDay = "5", deliveryDate = "21.01.2021", managerСomment = "skvjklsj", departmentComment = "svfkdsj" },
                new OrderBody { group = "2.1.3-", article = "12345678", barcode = "100000121", name = "HJKKHKAflkamn", supplier = "OAO JJFJF",
                    goods = "1", average = "4", stockDay = "5", deliveryDate = "21.01.2021", managerСomment = "skvjklsj", departmentComment = "svfkdsj" },
                new OrderBody { group = "2.1.3.", article = "12345678", barcode = "100000121", name = "HJKKHKAflkamn", supplier = "OAO JJFJF",
                    goods = "1", average = "4", stockDay = "5", deliveryDate = "21.01.2021", managerСomment = "skvjklsj", departmentComment = "svfkdsj" },
                new OrderBody { group = "2.1.2-", article = "12345678", barcode = "100000121", name = "HJKKHKAflkamn", supplier = "OAO JJFJF",
                    goods = "1", average = "4", stockDay = "5", deliveryDate = "21.01.2021", managerСomment = "skvjklsj", departmentComment = "svfkdsj" },
                new OrderBody { group = "2.1.3-", article = "12345678", barcode = "100000121", name = "HJKKHKAflkamn", supplier = "OAO JJFJF",
                    goods = "1", average = "4", stockDay = "5", deliveryDate = "21.01.2021", managerСomment = "skvjklsj", departmentComment = "svfkdsj" },
                new OrderBody { group = "2.1.2-", article = "12345678", barcode = "100000121", name = "HJKKHKAflkamn", supplier = "OAO JJFJF",
                    goods = "1", average = "4", stockDay = "5", deliveryDate = "21.01.2021", managerСomment = "skvjklsj", departmentComment = "svfkdsj" },
                new OrderBody { group = "2.1.3-", article = "12345678", barcode = "100000121", name = "HJKKHKAflkamn", supplier = "OAO JJFJF",
                    goods = "1", average = "4", stockDay = "5", deliveryDate = "21.01.2021", managerСomment = "skvjklsj", departmentComment = "svfkdsj" },
                new OrderBody { group = "2.1.3.", article = "12345678", barcode = "100000121", name = "HJKKHKAflkamn", supplier = "OAO JJFJF",
                    goods = "1", average = "4", stockDay = "5", deliveryDate = "21.01.2021", managerСomment = "skvjklsj", departmentComment = "svfkdsj" },
                new OrderBody { group = "2.1.2-", article = "12345678", barcode = "100000121", name = "HJKKHKAflkamn", supplier = "OAO JJFJF",
                    goods = "1", average = "4", stockDay = "5", deliveryDate = "21.01.2021", managerСomment = "skvjklsj", departmentComment = "svfkdsj" },
                new OrderBody { group = "2.1.3-", article = "12345678", barcode = "100000121", name = "HJKKHKAflkamn", supplier = "OAO JJFJF",
                    goods = "1", average = "4", stockDay = "5", deliveryDate = "21.01.2021", managerСomment = "skvjklsj", departmentComment = "svfkdsj" },
                new OrderBody { group = "2.1.3.", article = "12345678", barcode = "100000121", name = "HJKKHKAflkamn", supplier = "OAO JJFJF",
                    goods = "1", average = "4", stockDay = "5", deliveryDate = "21.01.2021", managerСomment = "skvjklsj", departmentComment = "svfkdsj" },
                new OrderBody { group = "2.1.2-", article = "12345678", barcode = "100000121", name = "HJKKHKAflkamn", supplier = "OAO JJFJF",
                    goods = "1", average = "4", stockDay = "5", deliveryDate = "21.01.2021", managerСomment = "skvjklsj", departmentComment = "svfkdsj" },
                new OrderBody { group = "2.1.3-", article = "12345678", barcode = "100000121", name = "HJKKHKAflkamn", supplier = "OAO JJFJF",
                    goods = "1", average = "4", stockDay = "5", deliveryDate = "21.01.2021", managerСomment = "skvjklsj", departmentComment = "svfkdsj" },
                new OrderBody { group = "2.1.3.", article = "12345678", barcode = "100000121", name = "HJKKHKAflkamn", supplier = "OAO JJFJF",
                    goods = "1", average = "4", stockDay = "5", deliveryDate = "21.01.2021", managerСomment = "skvjklsj", departmentComment = "svfkdsj" },
                new OrderBody { group = "2.1.2-", article = "12345678", barcode = "100000121", name = "HJKKHKAflkamn", supplier = "OAO JJFJF",
                    goods = "1", average = "4", stockDay = "5", deliveryDate = "21.01.2021", managerСomment = "skvjklsj", departmentComment = "svfkdsj" },
                new OrderBody { group = "2.1.3-", article = "12345678", barcode = "100000121", name = "HJKKHKAflkamn", supplier = "OAO JJFJF",
                    goods = "1", average = "4", stockDay = "5", deliveryDate = "21.01.2021", managerСomment = "skvjklsj", departmentComment = "svfkdsj" },
            },
        };

        public async Task<Stream> GetFile(FullOrder fullOrder)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Заказ");


                var colA = worksheet.Column("A");
                colA.Width = 10;

                var colB = worksheet.Column("B");
                colB.Width = 10;

                var colC = worksheet.Column("C");
                colC.Width = 10;

                var colD = worksheet.Column("D");
                colD.Style.Alignment.WrapText = true;
                colD.Width = 50;
                var colE = worksheet.Column("E");
                colE.Style.Alignment.WrapText = true;
                colE.Width = 20;

                //Order
                worksheet.Cell("A1").SetValue("Заказ");
                worksheet.Cell("C1").SetValue(_fullOrder.order.id);


                worksheet.Cell("H1").SetValue("Дата");
                worksheet.Cell("I1").SetValue(_fullOrder.order.creationDate);


                worksheet.Cell("H2").SetValue("Статус");
                worksheet.Cell("I2").SetValue(_fullOrder.order.statusId);


                worksheet.Cell("A3").SetValue("Магазин");
                worksheet.Cell("C3").SetValue(_fullOrder.order.storeId);

                worksheet.Cell("H3").SetValue("Отдел");
                worksheet.Cell("I3").SetValue(_fullOrder.order.deparmentId);

                worksheet.Cell("A5").SetValue("Автор - " + _fullOrder.order.author);
                worksheet.Cell("A5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                worksheet.Cell("A1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("A1").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Range("A1:B2").Merge();

                worksheet.Cell("C1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("C1").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("C1").Style.Font.FontSize = 22;
                worksheet.Range("C1:G2").Merge();

                worksheet.Range("I1:K1").Row(1).Merge();

                worksheet.Range("I2:K2").Row(1).Merge();

                worksheet.Cell("A3").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("A3").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Range("A3:B4").Merge();

                worksheet.Cell("C3").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("C3").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Range("C3:G4").Merge();

                worksheet.Cell("H3").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("H3").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Range("H3:H4").Column(1).Merge();

                worksheet.Cell("I3").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("I3").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Range("I3:K4").Merge();

                worksheet.Range("A5:K5").Row(1).Merge();


                //OrderBody
                worksheet.Cell("A6").SetValue("Группа");
                worksheet.Cell("B6").SetValue("Артикул");
                worksheet.Cell("C6").SetValue("Штпихкод");
                worksheet.Cell("D6").SetValue("Наименование");
                worksheet.Cell("E6").SetValue("Поставщик");
                worksheet.Cell("F6").SetValue("Остаток");
                worksheet.Cell("G6").SetValue("Ср. реал");
                worksheet.Cell("H6").SetValue("Запас дней");
                worksheet.Cell("I6").SetValue("Дата пост.");
                worksheet.Cell("J6").SetValue("Ком. рук.");
                worksheet.Cell("K6").SetValue("Примечание");

                //FreezeRows
                worksheet.SheetView.FreezeRows(6);

                var currentRow = 7;
                foreach (var item in _fullOrder.orderBodies)
                {
                    worksheet.Cell(currentRow, 1).Value = item.group;
                    worksheet.Cell(currentRow, 1).DataType = XLDataType.Text;

                    worksheet.Cell(currentRow, 2).Value = item.article;
                    worksheet.Cell(currentRow, 2).DataType = XLDataType.Number;

                    worksheet.Cell(currentRow, 3).Value = item.barcode;
                    worksheet.Cell(currentRow, 3).DataType = XLDataType.Number;

                    worksheet.Cell(currentRow, 4).Value = item.name;
                    worksheet.Cell(currentRow, 4).DataType = XLDataType.Text;

                    worksheet.Cell(currentRow, 5).Value = item.supplier;
                    worksheet.Cell(currentRow, 5).DataType = XLDataType.Text;

                    worksheet.Cell(currentRow, 6).Value = item.goods;
                    worksheet.Cell(currentRow, 6).DataType = XLDataType.Number;

                    worksheet.Cell(currentRow, 7).Value = item.average;
                    worksheet.Cell(currentRow, 7).DataType = XLDataType.Number;

                    worksheet.Cell(currentRow, 8).Value = item.stockDay;
                    worksheet.Cell(currentRow, 8).DataType = XLDataType.Number;

                    worksheet.Cell(currentRow, 9).Value = new DateTime(2010, 9, 2);
                    worksheet.Cell(currentRow, 9).DataType = XLDataType.DateTime;

                    worksheet.Cell(currentRow, 10).Value = item.managerСomment;
                    worksheet.Cell(currentRow, 10).DataType = XLDataType.Text;

                    worksheet.Cell(currentRow, 11).Value = item.departmentComment;
                    worksheet.Cell(currentRow, 11).DataType = XLDataType.Text;

                    currentRow++;
                }

                worksheet.Range($"A1:K{currentRow - 1}").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                worksheet.Range($"A1:K{currentRow - 1}").Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                worksheet.Range($"A1:K{currentRow - 1}").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                worksheet.Range($"A1:K{currentRow - 1}").Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                worksheet.Range($"A1:K{currentRow - 1}").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                worksheet.Range($"A1:K{currentRow - 1}").Style.Border.TopBorder = XLBorderStyleValues.Thin;

                // Adjust column width
                worksheet.Columns(6, 11).AdjustToContents();

                worksheet.PageSetup.PrintAreas.Add($"A1:K{currentRow - 1}");
                worksheet.PageSetup.PageOrientation = XLPageOrientation.Landscape;

                worksheet.PageSetup.SetRowsToRepeatAtTop(6, 6);


                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    //var content = stream.ToArray();

                    using (FileStream file = new FileStream("output.xlsx", FileMode.OpenOrCreate, FileAccess.Write))
                    {
                        stream.WriteTo(file);
                        file.Close();
                        stream.Close();
                    }

                    return stream;
                }
            }
        }
    }
}