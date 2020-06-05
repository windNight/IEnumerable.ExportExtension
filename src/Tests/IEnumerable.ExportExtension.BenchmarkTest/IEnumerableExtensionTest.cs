using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BenchmarkDotNet.Attributes;
using System.Linq.ExportExtension;
using System.IO;

namespace IEnumerable.ExportExtension.BenchmarkTest
{
    public class IEnumerableExtensionTest
    {
        static string ExportsDir => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Exports");
        //static readonly List<int> TestLines = new List<int> { 1000, 10000, 65530, 100000 };
        static readonly List<int> TestLines = new List<int> { 1000 };
        [Benchmark]
        public void TestEPPlusExport()
        {
            //var prefix = DateTime.Now.ToString("yyyyMMddHHmm");
            //var dir = Path.Combine(ExportsDir, prefix);
            //if (!Directory.Exists(dir))
            //{
            //    Directory.CreateDirectory(dir);
            //}
            var mode = "Sync";
            foreach (var item in TestLines)
            {
                var list = GenerateTestData(item);
                var bytesOfTestDataUseEPPlus = GetBytesOfTestDataUseEPPlus(list);
              //  SaveBytesOfTestDataUseEPPlus(bytesOfTestDataUseEPPlus, item, mode, 1, dir);
               // SaveTestDataUseEPPlus(list, item, mode, 1, dir);
            }

        }

        /// <summary>
        /// Generate Test Data
        /// </summary>
        /// <param name="lines"></param>
        /// <param name="serialNumber"></param>
        /// <param name="mode"></param>
        /// <param name="loopIndex"></param>
        /// <returns></returns>
        static IEnumerable<TT> GenerateTestData(int lines)
        {
            var list = GeneratorIEnumerable(lines);
            return list;
        }


        static IEnumerable<TT> GeneratorIEnumerable(int count)
        {
            var i = 0;
            while (i < count)
            {
                i++;
                yield return new TT
                {
                    Property1 = $"Property1_{i}",
                    Property2 = $"Property2_{i}",
                    Property3 = $"Property3_{i}",
                    Property4 = $"Property4_{i}",
                    Property5 = $"Property5_{i}",
                    Property6 = $"Property6_{i}",
                    Property7 = $"Property7_{i}",
                    Property8 = $"Property8_{i}",
                    Property9 = $"Property9_{i}",
                    Property10 = $"Property10_{i}",
                };
            };
        }

        #region  ExportHelper.EPPlus   

        // Just Get bytes use ExportHelper.EPPlus       And Do TimeWatcher
        static byte[] GetBytesOfTestDataUseEPPlus(IEnumerable<TT> data)
        {
            var bytesOfTestDataUseEPPlus = data.GetExportBytesByEPPlus();
            return bytesOfTestDataUseEPPlus;
        }

        /// <summary>
        /// save bytesOfTestDataUseEPPlus
        /// </summary>
        /// <param name="bytesOfTestDataUseEPPlus"></param>
        /// <param name="lines"></param>
        /// <param name="serialNumber"></param>
        /// <param name="loopIndex"></param>
        /// <param name="mode"></param>
        /// <param name="dir"></param>
        static void SaveBytesOfTestDataUseEPPlus(byte[] bytesOfTestDataUseEPPlus, int lines, string mode, int loopIndex, string dir)
        {
            // save bytesOfTestDataUseEPPlus
            var fileName = $"{loopIndex}_SaveAsBytes_UseEPPlus_{DateTime.Now:yyyyMMddHHmmsss}_{lines}_{mode}.xls";
            var filePath = Path.Combine(dir, fileName);

            ExportHelper.SaveBytes(filePath, bytesOfTestDataUseEPPlus);

        }
        static void SaveTestDataUseEPPlus(IEnumerable<TT> data, int lines, string mode, int loopIndex, string dir)
        {
            var fileName = $"{loopIndex}_ExportByEPPlus_{lines}_{mode}";
            var filePath = Path.Combine(dir, $"{fileName}.xls");
            data.ToDataTable().ExportByEPPlus(filePath);
        }

        #endregion //end ExportHelper.EPPlus   

    }


    /// <summary>  </summary>
    internal class TT
    {
        public string Property1 { get; set; }
        public string Property2 { get; set; }
        public string Property3 { get; set; }
        public string Property4 { get; set; }
        public string Property5 { get; set; }
        public string Property6 { get; set; }
        public string Property7 { get; set; }
        public string Property8 { get; set; }
        public string Property9 { get; set; }
        public string Property10 { get; set; }
    }

    internal class TestLog
    {
        public int Index { get; set; }
        public string SerialNumber { get; set; }
        public int Count { get; set; }
        public string Operator { get; set; }
        public string ConsumeTime { get; set; }
        public string Mode { get; set; }
        public int ManagedThreadId { get; set; }
        public string Msg { get; set; }
    }

}
