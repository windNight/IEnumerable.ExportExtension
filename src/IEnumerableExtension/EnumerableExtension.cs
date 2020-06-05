using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace IEnumerableExtension
{
    /// <summary>
    /// </summary>
    public static class EnumerableExtension
    {
        /// <summary>
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data"></param>
        /// <param name="columnsDict"></param>
        /// <returns></returns>
        public static DataTable ToDataTable<T>(this IList<T> data, Dictionary<string, string> columnsDict = null)
        {
            var table = new DataTable();
            if (typeof(T).IsValueType || typeof(T) == typeof(string))
            {
                var dc = new DataColumn("Value");
                table.Columns.Add(dc);
                foreach (var item in data)
                {
                    var dr = table.NewRow();
                    dr[0] = item;
                    table.Rows.Add(dr);
                }
            }
            else
            {
                var props = TypeDescriptor.GetProperties(typeof(T));
                foreach (PropertyDescriptor prop in props)
                    if (columnsDict != null && columnsDict.ContainsKey(prop.Name))
                        table.Columns.Add(columnsDict[prop.Name],
                            Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
                    else
                        table.Columns.Add(prop.Name,
                            Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);

                foreach (var item in data)
                {
                    var row = table.NewRow();

                    foreach (PropertyDescriptor prop in props)
                        try
                        {
                            if (columnsDict != null && columnsDict.ContainsKey(prop.Name))
                                row[columnsDict[prop.Name]] = prop.GetValue(item) ?? DBNull.Value;
                            else
                                row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                        }
                        catch (Exception ex)
                        { 
                            if (columnsDict != null && columnsDict.ContainsKey(prop.Name))
                                row[columnsDict[prop.Name]] = DBNull.Value;
                            else
                                row[prop.Name] = DBNull.Value;
                        }

                    table.Rows.Add(row);
                }
            }

            return table;
        }

        /// <summary>
        /// </summary>
        /// <param name="sourceTable"></param>
        /// <returns></returns>
        public static byte[] ExportByEPPlus(this DataTable sourceTable)
        {
            using (var pck = new ExcelPackage())
            {
                var sheetName = string.IsNullOrEmpty(sourceTable.TableName) ? "Sheet1" : sourceTable.TableName;
                var ws = pck.Workbook.Worksheets.Add(sheetName);

                ws.Cells["A1"].LoadFromDataTable(sourceTable, true);

                var borderStyle = ExcelBorderStyle.Thin;
                var borderColor = Color.FromArgb(155, 155, 155);

                using (var rng = ws.Cells[1, 1, sourceTable.Rows.Count + 1, sourceTable.Columns.Count])
                {
                    rng.Style.Font.Name = "宋体";
                    rng.Style.Font.Size = 10;
                    rng.Style.Fill.PatternType = ExcelFillStyle.Solid; //Set Pattern for the background to Solid
                    rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 255, 255));

                    rng.Style.Border.Top.Style = borderStyle;
                    rng.Style.Border.Top.Color.SetColor(borderColor);

                    rng.Style.Border.Bottom.Style = borderStyle;
                    rng.Style.Border.Bottom.Color.SetColor(borderColor);

                    rng.Style.Border.Right.Style = borderStyle;
                    rng.Style.Border.Right.Color.SetColor(borderColor);
                }

                using (var rng = ws.Cells[1, 1, 1, sourceTable.Columns.Count])
                {
                    rng.Style.Font.Bold = true;
                    rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(234, 241, 246)); //Set color to dark blue
                    rng.Style.Font.Color.SetColor(Color.FromArgb(51, 51, 51));
                }

                var bytes = pck.GetAsByteArray();
                return bytes;
            }
        }

        /// <summary>
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dt"></param>
        /// <returns></returns>
        public static IEnumerable<T> ToEnumerable<T>(this DataTable dt) where T : class, new()
        {
            var propertyInfos = typeof(T).GetProperties();

            foreach (DataRow row in dt.Rows)
            {
                var t = new T();
                foreach (var p in propertyInfos)
                    if (dt.Columns.IndexOf(p.Name) != -1 && row[p.Name] != DBNull.Value)
                        p.SetValue(t, row[p.Name], null);
                yield return t;
            }
        }
    }
}