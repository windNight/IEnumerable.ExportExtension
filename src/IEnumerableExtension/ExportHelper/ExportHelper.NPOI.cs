#if !NET35
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;

namespace System.Linq.ExportExtension
{
    public static partial class ExportHelper
    {
        /// <summary>
        /// GetExportBytesByNPOI
        /// </summary>
        /// <param name="data"></param>
        /// <remarks> 
        ///  Maximum number of columns in a worksheet (16384). 
        ///  Maximum number of rows in a worksheet (1048576).
        /// </remarks>
        /// <param name="extension">only support xlsx|xls </param>
        /// <returns></returns>
        public static byte[] GetExportBytesByNPOI<T>(this IEnumerable<T> data, string extension)
        {
            return GetExportBytesByNPOI(data.ToDataTable(), extension);
        }

        /// <summary>
        /// GetExportInfoByNPOI
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="extension">only support xlsx|xls </param>
        /// <returns></returns>
        public static byte[] GetExportBytesByNPOI(this DataTable dt, string extension)
        {

            IWorkbook workbook;

            if (extension == "xlsx")
            {
                workbook = new XSSFWorkbook();
            }
            else if (extension == "xls")
            {
                workbook = new HSSFWorkbook();
            }
            else
            {
                throw new Exception("This format is not supported");
            }

            ISheet sheet1 = workbook.CreateSheet("Sheet 1");

            //make a header row
            IRow row1 = sheet1.CreateRow(0);

            for (int j = 0; j < dt.Columns.Count; j++)
            {

                ICell cell = row1.CreateCell(j);
                String columnName = dt.Columns[j].ToString();
                cell.SetCellValue(columnName);
            }

            //loops through data
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                IRow row = sheet1.CreateRow(i + 1);
                for (int j = 0; j < dt.Columns.Count; j++)
                {

                    ICell cell = row.CreateCell(j);
                    String columnName = dt.Columns[j].ToString();
                    cell.SetCellValue(dt.Rows[i][columnName].ToString());
                }
            }

            using (var exportData = new MemoryStream())
            {
                workbook.Write(exportData);
                return exportData.ToArray();
            }
        }

        /// <summary>
        /// ExportByNPOI
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="extension">only support xlsx|xls </param>
        /// <param name="filePath"></param>
        public static void ExportByNPOI(this DataTable dt, string extension, string filePath)
        {
            var bytes = GetExportBytesByNPOI(dt, extension);
            File.WriteAllBytes(filePath, bytes);
        }


        /// <summary>
        /// 将excel中的数据导入到DataTable中
        /// </summary>
        /// <param name="fileName">excel工作薄sheet的名称</param>
        /// <param name="sheetName">excel工作薄sheet的名称</param>
        /// <param name="isFirstRowColumn">第一行是否是DataTable的列名</param>
        /// <returns>返回的DataTable</returns>
        public static DataTable GetDataTableFromFileByNPOI(string fileName, string sheetName = "Sheet1", bool isFirstRowColumn = true)
        {
            IWorkbook workbook;
            ISheet sheet = null;
            DataTable data = new DataTable();
            int startRow = 0;
            try
            {
                var fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
                if (fileName.IndexOf(".xlsx") > 0) // 2007版本
                {
                    workbook = new XSSFWorkbook(fs);
                }
                else if (fileName.IndexOf(".xls") > 0) // 2003版本
                {
                    workbook = new HSSFWorkbook(fs);
                }
                else
                {
                    throw new Exception("This format is not supported");
                }
                if (sheetName != null)
                {
                    sheet = workbook.GetSheet(sheetName);
                    if (sheet == null) //如果没有找到指定的sheetName对应的sheet，则尝试获取第一个sheet
                    {
                        sheet = workbook.GetSheetAt(0);
                    }
                }
                else
                {
                    sheet = workbook.GetSheetAt(0);
                }
                if (sheet != null)
                {
                    IRow firstRow = sheet.GetRow(0);
                    int cellCount = firstRow.LastCellNum; //一行最后一个cell的编号 即总的列数

                    if (isFirstRowColumn)
                    {
                        for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                        {
                            ICell cell = firstRow.GetCell(i);
                            if (cell != null)
                            {
                                string cellValue = cell.StringCellValue;
                                if (cellValue != null)
                                {
                                    DataColumn column = new DataColumn(cellValue);
                                    data.Columns.Add(column);
                                }
                            }
                        }
                        startRow = sheet.FirstRowNum + 1;
                    }
                    else
                    {
                        startRow = sheet.FirstRowNum;
                    }

                    //最后一列的标号
                    int rowCount = sheet.LastRowNum;
                    for (int i = startRow; i <= rowCount; ++i)
                    {
                        IRow row = sheet.GetRow(i);
                        if (row == null) continue; //没有数据的行默认是null　　　　　　　

                        DataRow dataRow = data.NewRow();
                        for (int j = row.FirstCellNum; j < cellCount; ++j)
                        {
                            if (row.GetCell(j) != null) //同理，没有数据的单元格都默认是null
                                dataRow[j] = row.GetCell(j).ToString();
                        }
                        data.Rows.Add(dataRow);
                    }
                }

                return data;
            }
            catch (Exception ex)
            {
                throw ex;
                //Console.WriteLine("Exception: " + ex.Message);
                // return null;
            }
        }

    }
}
#endif