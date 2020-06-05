using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace System.Linq.ExportExtension
{

    /// <summary>
    /// 
    /// </summary>
    public static partial class ExportHelper
    {
        /// <summary>
        /// Save Bytes to localFile
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="bytes"></param>
        public static void SaveBytes(string filePath, byte[] bytes)
        {
            File.WriteAllBytes(filePath, bytes);
        }


    }
}