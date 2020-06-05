
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace System.Linq.ExportExtension 
{
    public static partial class ExportHelper  
    {
        /// <summary>
        /// GetExportXDocument
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data"></param>
        /// <param name="encode"> default is UTF-8 </param>
        /// <returns></returns>
        public static byte[] GetExportXDocumentInfo<T>(this IEnumerable<T> data, string encode = "UTF-8")
        {
            var xds = GetExportXDocument(data);

            var settings = new XmlWriterSettings { OmitXmlDeclaration = true, Encoding = Encoding.GetEncoding(encode) };
            using (var memoryStream = new MemoryStream())
            using (var xmlWriter = XmlWriter.Create(memoryStream, settings))
            {
                xds.WriteTo(xmlWriter);
                xmlWriter.Flush();
                return memoryStream.ToArray();
            }
        }

        /// <summary>
        /// GetExportXDocument
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data"></param>
        /// <returns></returns>
        public static XDocument GetExportXDocument<T>(this IEnumerable<T> data)
        {
            var xds = JsonConvert.DeserializeXNode("{\"Info\":" + JsonConvert.SerializeObject(data) + "}", "Root");
            return xds;
        }

        /// <summary>
        /// ExportByXDocumentSave
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data"></param>
        /// <param name="filePath"></param>
        public static void ExportByXDocumentSave<T>(this IEnumerable<T> data, string filePath)
        {
            var xds = GetExportXDocument(data);
            ExportByXDocumentSave(xds, filePath);
        }

        /// <summary>
        /// ExportByXDocumentSave
        /// </summary>
        /// <param name="xds"></param>
        /// <param name="filePath"></param>
        public static void ExportByXDocumentSave(this XDocument xds, string filePath)
        {
            xds.Save(filePath);
        }


    }
}