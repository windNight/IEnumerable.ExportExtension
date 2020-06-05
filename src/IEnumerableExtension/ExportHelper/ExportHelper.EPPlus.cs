using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;

namespace System.Linq.ExportExtension
{
    public static partial class ExportHelper
    {
        /// <summary>
        /// GetExportBytesByEPPlus
        /// </summary>
        /// <param name="data"></param>
        /// <remarks> 
        ///  Maximum number of columns in a worksheet (16384). 
        ///  Maximum number of rows in a worksheet (1048576).
        /// </remarks>
        /// <returns></returns>
        public static byte[] GetExportBytesByEPPlus<T>(this IEnumerable<T> data)
        {
            return data.ToDataTable().GetExportBytesByEPPlus();
        }

        /// <summary>
        /// GetExportBytesByEPPlus
        /// </summary>
        /// <param name="dt"></param>
        /// <remarks> 
        ///  Maximum number of columns in a worksheet (16384). 
        ///  Maximum number of rows in a worksheet (1048576).
        /// </remarks>
        /// <returns></returns>
        public static byte[] GetExportBytesByEPPlus(this DataTable dt)
        {
            using (ExcelPackage pck = new ExcelPackage())
            {
                string sheetName = string.IsNullOrEmpty(dt.TableName) ? "Sheet1" : dt.TableName;
                ExcelWorksheet ws = pck.Workbook.Worksheets.Add(sheetName);

                ws.Cells["A1"].LoadFromDataTable(dt, true);

                ExcelBorderStyle borderStyle = ExcelBorderStyle.Thin;
                Color borderColor = Color.FromArgb(155, 155, 155);

                using (ExcelRange rng = ws.Cells[1, 1, dt.Rows.Count + 1, dt.Columns.Count])
                {
                    rng.Style.Font.Name = "宋体";
                    rng.Style.Font.Size = 10;
                    rng.Style.Fill.PatternType = ExcelFillStyle.Solid;                      //Set Pattern for the background to Solid
                    rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 255, 255));

                    rng.Style.Border.Top.Style = borderStyle;
                    rng.Style.Border.Top.Color.SetColor(borderColor);

                    rng.Style.Border.Bottom.Style = borderStyle;
                    rng.Style.Border.Bottom.Color.SetColor(borderColor);

                    rng.Style.Border.Right.Style = borderStyle;
                    rng.Style.Border.Right.Color.SetColor(borderColor);
                }

                using (ExcelRange rng = ws.Cells[1, 1, 1, dt.Columns.Count])
                {
                    rng.Style.Font.Bold = true;
                    rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(234, 241, 246));  //Set color to dark blue
                    rng.Style.Font.Color.SetColor(Color.FromArgb(51, 51, 51));
                }

                return pck.GetAsByteArray();
            }
        }

        /// <summary>
        /// ExportByEPPlus
        /// </summary>
        /// <remarks> 
        ///  Maximum number of columns in a worksheet (16384). 
        ///  Maximum number of rows in a worksheet (1048576).
        /// </remarks>
        /// <param name="dt"></param>
        /// <param name="filePath"></param>
        public static void ExportByEPPlus(this DataTable dt, string filePath)
        {
            var bytes = GetExportBytesByEPPlus(dt);
            using (var exportData = new MemoryStream())
            {
                SaveBytes(filePath, bytes);
            }
        }

        public static object GetDataTableFromFileByEPPlus(string filePath)
        {
            using (ExcelPackage package = new ExcelPackage(new FileStream(filePath, FileMode.Open)))
            {
                for (int i = 1; i <= package.Workbook.Worksheets.Count; ++i)
                {
                    ExcelWorksheet sheet = package.Workbook.Worksheets[i];
                    for (int j = sheet.Dimension.Start.Column, k = sheet.Dimension.End.Column; j <= k; j++)
                    {
                        for (int m = sheet.Dimension.Start.Row, n = sheet.Dimension.End.Row; m <= n; m++)
                        {
                            var str = sheet.GetValue(m, j);
                            if (str != null)
                            {
                                // do something
                            }
                        }
                    }
                }
            }
            return null;
            // Func<CustomAttributeData, bool> columnOnly = y => y.AttributeType == typeof(ExcelColumn);
            //var columns = typeof(T)
            //    .GetProperties()
            //    .Where(x => x.CustomAttributes.Any(columnOnly))
            //    .Select(p => new
            //    {
            //        Property = p,
            //        Column = p.GetCustomAttributes<ExcelColumn>().First().ColumnName 
            //    }).ToList();

            //var rows = worksheet.Cells
            //    .Select(cell => cell.Start.Row)
            //    .Distinct()
            //    .OrderBy(x => x);

            //var collection = rows.Skip(1)
            //    .Select(row =>
            //    {
            //        var tnew = new T();
            //        columns.ForEach(col =>
            //        {
            //            var val = worksheet.Cells[row, GetColumnByName(worksheet, col.Column)];
            //            if (val.Value == null)
            //            {
            //                col.Property.SetValue(tnew, null);
            //                return;
            //            }
            //            // 如果Person类的对应字段是int的，该怎么怎么做……
            //            if (col.Property.PropertyType == typeof(int))
            //            {
            //                col.Property.SetValue(tnew, val.GetValue<int>());
            //                return;
            //            }
            //            // 如果Person类的对应字段是double的，该怎么怎么做……
            //            if (col.Property.PropertyType == typeof(double))
            //            {
            //                col.Property.SetValue(tnew, val.GetValue<double>());
            //                return;
            //            }
            //            // 如果Person类的对应字段是DateTime?的，该怎么怎么做……
            //            if (col.Property.PropertyType == typeof(DateTime?))
            //            {
            //                col.Property.SetValue(tnew, val.GetValue<DateTime?>());
            //                return;
            //            }
            //            // 如果Person类的对应字段是DateTime的，该怎么怎么做……
            //            if (col.Property.PropertyType == typeof(DateTime))
            //            {
            //                col.Property.SetValue(tnew, val.GetValue<DateTime>());
            //                return;
            //            }
            //            // 如果Person类的对应字段是bool的，该怎么怎么做……
            //            if (col.Property.PropertyType == typeof(bool))
            //            {
            //                col.Property.SetValue(tnew, val.GetValue<bool>());
            //                return;
            //            }
            //            col.Property.SetValue(tnew, val.GetValue<string>());
            //        });

            //        return tnew;
            //    });
            //return collection; 
        }
    }
}