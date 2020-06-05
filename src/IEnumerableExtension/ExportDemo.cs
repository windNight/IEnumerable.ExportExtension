using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Reflection;
using System.Threading;
using System.IO;
using System.Linq;
using System.Data;
using System.ComponentModel;
using System.Linq.ExportExtension;

namespace IEnumerableExtension.Demos
{
    /// <summary>
    ///  Export File ExportHelper.CsvString  And Do TimeWatcher
    ///  Export File ExportHelper.EPPlus     And Do TimeWatcher
    ///  Export File ExportHelper.NPOI       And Do TimeWatcher
    ///  Export File ExportHelper.XDocument  And Do TimeWatcher
    /// </summary>
    public class ExportDemo
    {
        static List<TestLog> logs = new List<TestLog>();
        static readonly List<int> TestLines = new List<int> { 1000, 10000, 65530, 100000 };
        static string LogsDir => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Logs");
        static string ExportsDir => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Exports");

        /// <summary>
        /// RunTest With LoopCount
        /// </summary>
        /// <param name="count">default is 1.</param>
        public static void RunTest(int count = 1)
        {
            logs = new List<TestLog>();
            var prefix = DateTime.Now.ToString("yyyyMMddHHmm");
            for (int i = 0; i < count; i++)
            {
                LoopTestSync(i, prefix);
                Console.WriteLine($"Test Finish Loop {i}");
            }
            Console.WriteLine("Test Finish ALL");
            SaveTestLogs(prefix);
            Console.WriteLine("SaveTestLogs Finish");
        }

        static void LoopTestSync(int loopIndex, string prefix)
        {
            prefix = $"{prefix}_Sync";
            foreach (var item in TestLines)
            {
                Console.WriteLine($"==TestSync_{item}Begin==".PadLeft(120, '=').PadRight(235, '='));
                TimeWatcher(() => TestSync(item, prefix, loopIndex), item, loopIndex, $"TestSync(1000,{prefix})", $"TestCount_{item}", "Sync");
                Console.WriteLine($"==TestSync_{item} End ==".PadLeft(120, '=').PadRight(235, '='));
            }
            Console.WriteLine($"LoopTestSync Finish");

        }

        static void TestSync(int lines, string prefix, int loopIndex)
        {
            var dir = Path.Combine(ExportsDir, prefix);
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }
            var serialNumber = GetGuid();

            var mode = "Sync";
            var linesStr = lines.ToString().PadLeft(8, '0');

            //Generate Test Data
            var list = GenerateTestData(lines, serialNumber, mode, loopIndex);

            #region  ExportHelper.CsvString 

            TimeWatcher(() =>
            {
                // Just Get bytes use ExportHelper.CsvString And Do TimeWatcher
                var bytesOfTestDataUseCsvString = GetBytesOfTestDataUseCsvString(list, lines, serialNumber, mode, loopIndex);
                SaveBytesOfTestDataUseCsvString(bytesOfTestDataUseCsvString, lines, serialNumber, mode, loopIndex, dir);
            }, lines, loopIndex, serialNumber, $"ExportHelper.CsvString_{linesStr}_Total", $"ExportHelper.CsvString_{linesStr}_Total", mode);

            SaveTestDataUseCsvString(list, lines, serialNumber, mode, loopIndex, dir);

            #endregion //end  ExportHelper.CsvString 

            #region ExportHelper.EPPlus
            TimeWatcher(() =>
            {
                // Just Get bytes use ExportHelper.EPPlus And Do TimeWatcher
                var bytesOfTestDataUseEPPlus = GetBytesOfTestDataUseEPPlus(list, lines, serialNumber, mode, loopIndex);
                SaveBytesOfTestDataUseEPPlus(bytesOfTestDataUseEPPlus, lines, serialNumber, mode, loopIndex, dir);
            }, lines, loopIndex, serialNumber, $"ExportHelper.EPPlus_{linesStr}_Total", $"ExportHelper.EPPlus_{linesStr}_Total", mode);


            SaveTestDataUseEPPlus(list, lines, serialNumber, mode, loopIndex, dir);

            #endregion //end ExportHelper.EPPlus

            #region ExportHelper.XDocument
            TimeWatcher(() =>
            {
                // Just Get bytes use ExportHelper.XDocument And Do TimeWatcher
                var bytesOfTestDataUseXDocument = GetBytesOfTestDataUseXDocument(list, lines, serialNumber, mode, loopIndex);
                SaveBytesOfTestDataUseXDocument(bytesOfTestDataUseXDocument, lines, serialNumber, mode, loopIndex, dir);
            }, lines, loopIndex, serialNumber, $"ExportHelper.XDocument_{linesStr}_Total", $"ExportHelper.XDocument_{linesStr}_Total", mode);
            SaveTestDataUseXDocument(list, lines, serialNumber, mode, loopIndex, dir);
            #endregion //end  ExportHelper.XDocument

            #region Test NPOI above net35
#if !NET35
            if (lines < 65536)
            {
                TimeWatcher(() =>
                {
                    // Just Get bytes use ExportHelper.NPOI And Do TimeWatcher
                    var bytesOfTestDataUseNPOI = GetBytesOfTestDataUseNPOI(list, lines, serialNumber, mode, loopIndex);
                    SaveBytesOfTestDataUseXDocument(bytesOfTestDataUseNPOI, lines, serialNumber, mode, loopIndex, dir);
                }, lines, loopIndex, serialNumber, $"ExportHelper.NPOI_{linesStr}_Total", $"ExportHelper.NPOI_{linesStr}_Total", mode);

                SaveTestDataUseNPOI(list, lines, serialNumber, mode, loopIndex, dir);
            }
#endif
            #endregion


        }

        /// <summary>
        /// Generate Test Data
        /// </summary>
        /// <param name="lines"></param>
        /// <param name="serialNumber"></param>
        /// <param name="mode"></param>
        /// <param name="loopIndex"></param>
        /// <returns></returns>
        static IEnumerable<TT> GenerateTestData(int lines, string serialNumber, string mode, int loopIndex)
        {
            var linesStr = lines.ToString().PadLeft(8, '0');
            var msg = $"LoopIndex {loopIndex},GeneratorIEnumerable_{linesStr} ";
            var list = TimeWatcher(() => GeneratorIEnumerable(lines), lines, loopIndex, serialNumber, $"{nameof(GeneratorIEnumerable)}{linesStr}", msg, mode);
            return list;
        }

        #region  ExportHelper.CsvString   

        /// <summary>
        /// Just Get bytes use ExportHelper.CsvString And Do TimeWatcher
        /// </summary>
        /// <param name="data"></param>
        /// <param name="lines"></param>
        /// <param name="serialNumber"></param>
        /// <param name="mode"></param>
        /// <param name="loopIndex"></param>
        /// <returns></returns>
        static byte[] GetBytesOfTestDataUseCsvString(IEnumerable<TT> data, int lines, string serialNumber, string mode, int loopIndex)
        {
            var linesStr = lines.ToString().PadLeft(8, '0');
            var msg = $"LoopIndex {loopIndex},{nameof(GetBytesOfTestDataUseCsvString)}_{linesStr} in {mode} way";
            var bytesOfTestDataUseCsvString = TimeWatcher(() => data.GenerateCsvBytes(), lines, loopIndex, serialNumber, $"{nameof(GetBytesOfTestDataUseCsvString)}_{linesStr}", msg, mode);
            return bytesOfTestDataUseCsvString;
        }

        /// <summary>
        /// save bytesOfTestDataUseCsvString
        /// </summary>
        /// <param name="bytesOfTestDataUseCsvString"></param>
        /// <param name="lines"></param>
        /// <param name="serialNumber"></param>
        /// <param name="mode"></param>
        /// <param name="loopIndex"></param>
        /// <param name="dir"></param>
        static void SaveBytesOfTestDataUseCsvString(byte[] bytesOfTestDataUseCsvString, int lines, string serialNumber, string mode, int loopIndex, string dir)
        {
            var linesStr = lines.ToString().PadLeft(8, '0');
            var msg = $"LoopIndex {loopIndex},{nameof(SaveBytesOfTestDataUseCsvString)}_{linesStr} in {mode} way . File length is {bytesOfTestDataUseCsvString.Length}.";
            var fileName = $"{loopIndex}_SaveAsBytes_UseCsvString_{DateTime.Now.ToString("yyyyMMddHHmmsss")}_{lines}_{mode}.csv";
            var filePath = Path.Combine(dir, fileName);
            TimeWatcher(() => ExportHelper.SaveBytes(filePath, bytesOfTestDataUseCsvString), lines, loopIndex, serialNumber, $"{nameof(SaveBytesOfTestDataUseCsvString)}_{linesStr}", msg, mode);
        }

        /// <summary>
        /// Just Get bytes use ExportHelper.CsvString And Do TimeWatcher
        /// </summary>
        /// <param name="data"></param>
        /// <param name="lines"></param>
        /// <param name="serialNumber"></param>
        /// <param name="mode"></param>
        /// <param name="loopIndex"></param>
        /// <param name="dir"></param>
        /// <returns></returns>
        static void SaveTestDataUseCsvString(IEnumerable<TT> data, int lines, string serialNumber, string mode, int loopIndex, string dir)
        {
            var linesStr = lines.ToString().PadLeft(8, '0');
            var msg = $"LoopIndex {loopIndex},{nameof(SaveTestDataUseCsvString)}_{linesStr} in {mode} way";
            var fileName = $"{loopIndex}_ExportByCsvString_{lines}_{mode}";
            var filePath = Path.Combine(dir, $"{fileName}.csv");
            TimeWatcher(() => data.ExportByCsvString(filePath),
                lines, loopIndex, serialNumber, fileName, msg, mode);
        }

        #endregion

        #region  ExportHelper.EPPlus   

        // Just Get bytes use ExportHelper.EPPlus       And Do TimeWatcher
        static byte[] GetBytesOfTestDataUseEPPlus(IEnumerable<TT> data, int lines, string serialNumber, string mode, int loopIndex)
        {
            // Just Get bytes use ExportHelper.EPPlus And Do TimeWatcher
            var linesStr = lines.ToString().PadLeft(8, '0');
            var msg = $"LoopIndex {loopIndex},{nameof(GetBytesOfTestDataUseEPPlus)}_{linesStr} in {mode} way";
            var bytesOfTestDataUseEPPlus = TimeWatcher(() => data.GetExportBytesByEPPlus(), lines, loopIndex, serialNumber, $"{nameof(GetBytesOfTestDataUseEPPlus)}_{linesStr}", msg, mode);

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
        static void SaveBytesOfTestDataUseEPPlus(byte[] bytesOfTestDataUseEPPlus, int lines, string serialNumber, string mode, int loopIndex, string dir)
        {
            // save bytesOfTestDataUseEPPlus
            var linesStr = lines.ToString().PadLeft(8, '0');

            var msg = $"LoopIndex {loopIndex},{nameof(SaveBytesOfTestDataUseEPPlus)}_{linesStr} in {mode} way. File length is {bytesOfTestDataUseEPPlus.Length}.";
            var fileName = $"{loopIndex}_SaveAsBytes_UseEPPlus_{DateTime.Now.ToString("yyyyMMddHHmmsss")}_{lines}_{mode}.xls";
            var filePath = Path.Combine(dir, fileName);

            TimeWatcher(() => ExportHelper.SaveBytes(filePath, bytesOfTestDataUseEPPlus), lines, loopIndex, serialNumber, $"{nameof(SaveBytesOfTestDataUseEPPlus)}_{linesStr}", msg, mode);

        }
        static void SaveTestDataUseEPPlus(IEnumerable<TT> data, int lines, string serialNumber, string mode, int loopIndex, string dir)
        {
            var linesStr = lines.ToString().PadLeft(8, '0');

            var msg = $"LoopIndex {loopIndex},{nameof(SaveTestDataUseEPPlus)}_{linesStr} in {mode} way. ";
            var fileName = $"{loopIndex}_ExportByEPPlus_{lines}_{mode}";
            var filePath = Path.Combine(dir, $"{fileName}.xls");

            TimeWatcher(() => data.ToDataTable().ExportByEPPlus(filePath),
                lines, loopIndex, serialNumber, fileName, msg, mode
                );
        }

        #endregion

        #region  ExportHelper.XDocument   

        /// <summary>
        /// Just Get bytes use ExportHelper.XDocument And Do TimeWatcher
        /// </summary>
        /// <param name="data"></param>
        /// <param name="lines"></param>
        /// <param name="serialNumber"></param>
        /// <param name="mode"></param>
        /// <param name="loopIndex"></param>
        /// <returns></returns>
        static byte[] GetBytesOfTestDataUseXDocument(IEnumerable<TT> data, int lines, string serialNumber, string mode, int loopIndex)
        {
            var linesStr = lines.ToString().PadLeft(8, '0');
            // Just Get bytes use ExportHelper.XDocument And Do TimeWatcher
            var msg = $"LoopIndex {loopIndex},{nameof(GetBytesOfTestDataUseXDocument)}_{linesStr} in {mode} way";
            var bytesOfTestDataUseXDocument = TimeWatcher(() => data.GetExportXDocumentInfo(), lines, loopIndex, serialNumber, $"{nameof(GetBytesOfTestDataUseXDocument)}{linesStr}", msg, mode);

            return bytesOfTestDataUseXDocument;
        }

        /// <summary>
        /// save bytesOfTestDataUseXDocument
        /// </summary>
        /// <param name="bytesOfTestDataUseXDocument"></param>
        /// <param name="lines"></param>
        /// <param name="serialNumber"></param>
        /// <param name="mode"></param>
        /// <param name="loopIndex"></param>
        static void SaveBytesOfTestDataUseXDocument(byte[] bytesOfTestDataUseXDocument, int lines, string serialNumber, string mode, int loopIndex, string dir)
        {
            var linesStr = lines.ToString().PadLeft(8, '0');
            // save bytesOfTestDataUseXDocument
            var msg = $"LoopIndex {loopIndex},{nameof(SaveBytesOfTestDataUseXDocument)}_{linesStr} in {mode} way . File length is {bytesOfTestDataUseXDocument.Length}.";
            var fileName = $"{loopIndex}_SaveAsBytes_UseXDocument_{DateTime.Now.ToString("yyyyMMddHHmmsss")}_{lines}_{mode}.xls";
            var filePath = Path.Combine(dir, fileName);
            TimeWatcher(() => ExportHelper.SaveBytes(filePath, bytesOfTestDataUseXDocument), lines, loopIndex, serialNumber, $"{nameof(SaveBytesOfTestDataUseXDocument)}_{linesStr}", msg, mode);

        }

        static void SaveTestDataUseXDocument(IEnumerable<TT> data, int lines, string serialNumber, string mode, int loopIndex, string dir)
        {
            var linesStr = lines.ToString().PadLeft(8, '0');

            var msg = $"LoopIndex {loopIndex},{nameof(SaveTestDataUseXDocument)}_{linesStr} in {mode} way. ";
            var fileName = $"{loopIndex}_ExportByXDocumentSave_{lines}_{mode}";
            var filePath = Path.Combine(dir, $"{fileName}.xls");

            TimeWatcher(() => data.ExportByXDocumentSave(filePath),
                lines, loopIndex, serialNumber, fileName, msg, mode
                );
        }

        #endregion

        #region  ExportHelper.NPOI   
#if !NET35
        #region Test NPOI above net35
        /// <summary>
        /// Just Get bytes use ExportHelper.NPOI And Do TimeWatcher
        /// </summary>
        /// <param name="data"></param>
        /// <param name="lines"></param>
        /// <param name="serialNumber"></param>
        /// <param name="mode"></param>
        /// <param name="loopIndex"></param>
        /// <returns></returns>
        static byte[] GetBytesOfTestDataUseNPOI(IEnumerable<TT> data, int lines, string serialNumber, string mode, int loopIndex)
        {
            var linesStr = lines.ToString().PadLeft(8, '0');
            var msg = $"LoopIndex {loopIndex},{nameof(GetBytesOfTestDataUseNPOI)}_{linesStr} in {mode} way";

            var bytesOfTestDataUseNPOI = TimeWatcher(() => data.GetExportBytesByNPOI("xls"), lines, loopIndex, serialNumber, $"{nameof(GetBytesOfTestDataUseNPOI)}_{linesStr}", msg, mode);

            return bytesOfTestDataUseNPOI;
        }

        /// <summary>
        /// save bytesOfTestDataUseXDocument
        /// </summary>
        /// <param name="bytesOfTestDataUseNPOI"></param>
        /// <param name="lines"></param>
        /// <param name="serialNumber"></param>
        /// <param name="mode"></param>
        /// <param name="loopIndex"></param>
        public static void SaveBytesOfTestDataUseNPOI(byte[] bytesOfTestDataUseNPOI, int lines, string serialNumber, string mode, int loopIndex, string dir)
        {
            var linesStr = lines.ToString().PadLeft(8, '0');
            // save bytesOfTestDataUseXDocument
            var msg = $"LoopIndex {loopIndex},{nameof(SaveBytesOfTestDataUseNPOI)}_{linesStr} in {mode} way. File length is {bytesOfTestDataUseNPOI.Length}.";
            var fileName = $"{loopIndex}_SaveAsBytes_UseNPOI_{DateTime.Now.ToString("yyyyMMddHHmmsss")}_{lines}_{mode}.xml";
            var filePath = Path.Combine(dir, fileName);
            TimeWatcher(() => ExportHelper.SaveBytes(filePath, bytesOfTestDataUseNPOI), lines, loopIndex, serialNumber, $"{nameof(SaveBytesOfTestDataUseNPOI)}_{linesStr}", msg, mode);

        }

        static void SaveTestDataUseNPOI(IEnumerable<TT> data, int lines, string serialNumber, string mode, int loopIndex, string dir)
        {
            var linesStr = lines.ToString().PadLeft(8, '0');

            var msg = $"LoopIndex {loopIndex},ExportByNPOI_{linesStr} in {mode} way. ";
            var fileName = $"{loopIndex}_ExportByNPOI_{lines}_{mode}";
            var filePath = Path.Combine(dir, $"{fileName}.xls");

            TimeWatcher(() => data.ToDataTable().ExportByNPOI("xls", filePath),
                lines, loopIndex, serialNumber, fileName, msg, mode
                );
        }
        #endregion
#endif
        #endregion

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

        #region Common
        static void SaveTestLogs(string prefix)
        {
            if (logs != null)
            {
                var dir = Path.Combine(LogsDir, prefix);
                if (!Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }
                foreach (var opt in logs.Select(m => m.Operator).Distinct())
                {
                    var optFileName = $"Logs_{opt}_{DateTime.Now:yyyyMMddHHmmsss}.xls";
                    var optFilePath = Path.Combine(dir, optFileName);
                    logs.Where(m => m.Operator == opt).OrderBy(m => m.Index).ExportByXDocumentSave(optFilePath);
                };

                var fileName = $"Logs_{DateTime.Now:yyyyMMddHHmmsss}.xls";
                var filePath = Path.Combine(dir, fileName);
                logs.ExportByXDocumentSave(filePath);
            }
        }


        static void TimeWatcher(Action action, int testCount, int testIndex, string serialNumber, string actionName, string logMsg, string mode = "")
        {
            var ticks = DateTime.Now.Ticks;
            action.Invoke();
            var consumeTime = TimeSpan.FromTicks(DateTime.Now.Ticks - ticks).TotalMilliseconds;
            var msg = $"{logMsg.PadRight(120, ' ')} consume {consumeTime.ToString("F4").PadLeft(12, ' ')} ms. mode is {mode.PadLeft(5, ' ')}. ManagedThread is { Thread.CurrentThread.ManagedThreadId.ToString().PadLeft(3, ' ')}|{Thread.CurrentThread.Name}";
            Console.WriteLine(msg);
            logs.Add(new TestLog
            {
                Index = testIndex,
                Count = testCount,
                Operator = actionName,
                ConsumeTime = consumeTime.ToString("F4"),
                SerialNumber = serialNumber,
                Mode = mode,
                ManagedThreadId = Thread.CurrentThread.ManagedThreadId,
                Msg = msg,
            });
        }

        static T TimeWatcher<T>(Func<T> action, int testCount, int testIndex, string serialNumber, string actionName, string logMsg, string mode = "")
        {
            var ticks = DateTime.Now.Ticks;
            var rlt = action.Invoke();
            var consumeTime = TimeSpan.FromTicks(DateTime.Now.Ticks - ticks).TotalMilliseconds;
            var msg = $"{logMsg.PadRight(120, ' ')} consume {consumeTime.ToString("F4").PadLeft(12, ' ')} ms. mode is {mode.PadLeft(5, ' ')}. ManagedThread is { Thread.CurrentThread.ManagedThreadId.ToString().PadLeft(3, ' ')}|{Thread.CurrentThread.Name}";
            Console.WriteLine(msg);

            logs.Add(new TestLog
            {
                Index = testIndex,
                Count = testCount,
                Operator = actionName,
                ConsumeTime = consumeTime.ToString("F4"),
                SerialNumber = serialNumber,
                Mode = mode,
                ManagedThreadId = Thread.CurrentThread.ManagedThreadId,
                Msg = msg,
            });
            return rlt;
        }

        static string GetGuid()
        {
            string strDateTimeNumber = DateTime.Now.ToString("yyyyMMddHHmmssms");
            string strRandomResult = NextRandom(1000, 1).ToString("D3");
            return string.Concat(strDateTimeNumber, strRandomResult);
        }

        static int NextRandom(int numSeeds, int length)
        {
            byte[] randomNumber = new byte[length];
            System.Security.Cryptography.RNGCryptoServiceProvider rng = new System.Security.Cryptography.RNGCryptoServiceProvider();
            rng.GetBytes(randomNumber);
            uint randomResult = 0x0;
            for (int i = 0; i < length; i++)
            {
                randomResult |= ((uint)randomNumber[i] << ((length - 1 - i) * 8));
            }
            return (int)(randomResult % numSeeds) + 1;
        }


        #endregion

    }

    /// <summary>
    /// 
    /// </summary>
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