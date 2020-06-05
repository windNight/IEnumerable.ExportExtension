using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;

namespace System.Linq.ExportExtension
{
    /// <summary> </summary>
    public static partial class ExportHelper
    {
        /// <summary>
        /// GenerateCsvBytes
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data"></param>
        /// <param name="propList">自定义属性列表</param>
        /// <param name="encode"> default is UTF-8 </param>
        /// <returns></returns>
        public static byte[] GenerateCsvBytes<T>(this IEnumerable<T> data, IEnumerable<string> propList = null, string encode = "UTF-8", bool needTitle = true)
        {
            var dataCsvString = GenerateCsvString(data, propList, needTitle);
            return Encoding.GetEncoding(encode).GetBytes(dataCsvString);
        }

        /// <summary>
        /// GenerateCsvString 
        /// split use ',' per Property 
        /// 
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="propList">自定义属性列表</param>
        /// <param name="data"></param>
        /// <returns></returns>
        public static string GenerateCsvString<T>(this IEnumerable<T> data, IEnumerable<string> propList = null, bool needTitle = true)
        {
            var str = new StringBuilder();
            if (needTitle)
            {
                if (propList == null || !propList.Any())
                    propList = GeneratePropertyNames<T>();
                str.AppendLine(string.Join(",", propList.ToArray()));
            }


            var valueList = GeneratePropertyCsvStrs(data);
            if (valueList != null)
                str.Append(string.Join("\r\n", valueList.ToArray()));
            return str.ToString();
        }

        /// <summary>
        ///  ExportByCsvString
        ///  split use ',' per Property 
        /// 
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data">  </param>
        /// <param name="filePath">the full filePath to save data  <paramref name="data"/>.
        ///  extension of the fileName suggest is '.csv'
        /// </param>
        public static void ExportByCsvString<T>(this IEnumerable<T> data, string filePath)
        {
            var csvStr = GenerateCsvString(data);
            File.WriteAllText(filePath, csvStr);
        }

        #region ==========Private========

        static IEnumerable<string> GeneratePropertyNames<T>()
        {
            PropertyDescriptorCollection props = TypeDescriptor.GetProperties(typeof(T));
            foreach (PropertyDescriptor prop in props)
            {
                yield return prop.Name;
            }
        }

        static IEnumerable<string> GeneratePropertyCsvStrs<T>(IEnumerable<T> data)
        {
            foreach (var item in data)
            {
                yield return string.Join(",", GeneratePropertyValues(item).ToArray());
            }
        }

        static IEnumerable<string> GeneratePropertyValues<T>(T data)
        {
            PropertyDescriptorCollection props = TypeDescriptor.GetProperties(typeof(T));
            foreach (PropertyDescriptor prop in props)
            {
                yield return prop.GetValue(data)?.ToString() ?? string.Empty;
            }
        }

        #endregion //end Private

    }
}