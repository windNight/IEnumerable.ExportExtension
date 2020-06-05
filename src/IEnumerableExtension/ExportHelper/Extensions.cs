using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Text;

namespace System.Linq.ExportExtension
{
    public static class DateTableExtensions
    {
        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data"></param>
        /// <param name="columnsDict"></param>
        /// <returns></returns>
        public static DataTable ToDataTable<T>(this IEnumerable<T> data, Dictionary<string, string> columnsDict = null)
        {
            DataTable table = new DataTable();
            if (typeof(T).IsValueType || typeof(T) == typeof(string))
            {

                DataColumn dc = new DataColumn("Value");
                table.Columns.Add(dc);
                foreach (T item in data)
                {
                    DataRow dr = table.NewRow();
                    dr[0] = item;
                    table.Rows.Add(dr);
                }
            }
            else
            {
                PropertyDescriptorCollection props = TypeDescriptor.GetProperties(typeof(T));
                foreach (PropertyDescriptor prop in props)
                {
                    if (columnsDict != null && columnsDict.ContainsKey(prop.Name))
                        table.Columns.Add(columnsDict[prop.Name], Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
                    else
                        table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
                }

                foreach (T item in data)
                {
                    DataRow row = table.NewRow();

                    foreach (PropertyDescriptor prop in props)
                    {
                        try
                        {
                            if (columnsDict != null && columnsDict.ContainsKey(prop.Name))
                                row[columnsDict[prop.Name]] = prop.GetValue(item) ?? DBNull.Value;
                            else
                                row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                        }
                        catch
                        {
                            if (columnsDict != null && columnsDict.ContainsKey(prop.Name))
                                row[columnsDict[prop.Name]] = DBNull.Value;
                            else
                                row[prop.Name] = DBNull.Value;
                        }
                    }

                    table.Rows.Add(row);
                }
            }

            return table;
        }
    }
}