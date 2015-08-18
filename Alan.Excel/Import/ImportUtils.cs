using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;

namespace Alan.Excel.Import
{
    public static class ImportUtils
    {

        /// <summary>
        /// 遍历所有Sheets
        /// </summary>
        /// <param name="fileFullPath">Excel文件绝对路径</param>
        /// <param name="callback">回调</param>
        public static void Sheets(string fileFullPath, Action<ExcelWorksheets> callback)
        {
            var fs = new FileInfo(fileFullPath);
            List<ExcelWorksheet> sheets;
            using (ExcelPackage package = new ExcelPackage(fs))
            {
                callback(package.Workbook.Worksheets);
            }
        }

        /// <summary>
        /// 获取指定的Sheet
        /// </summary>
        /// <param name="fileFullPath">Excel文件绝对路径</param>
        /// <param name="sheetName">Excel Sheet名字</param>
        /// <param name="callback">回调</param>
        public static void Sheet(string fileFullPath, string sheetName, Action<ExcelWorksheet> callback)
        {
            Sheets(fileFullPath, sheets =>
            {
                var sheet = sheets[sheetName];
                if (sheet == null) return;
                callback(sheet);
            });
        }

        /// <summary>
        /// 获取指定的Sheet
        /// </summary>
        /// <param name="fileFullPath">Excel文件绝对路径</param>
        /// <param name="index">Excel Sheet索引</param>
        /// <param name="callback">回调</param>
        public static void Sheet(string fileFullPath, int index, Action<ExcelWorksheet> callback)
        {
            Sheets(fileFullPath, sheets =>
            {
                var sheet = sheets[index];
                if (sheet == null) return;
                callback(sheet);
            });
        }
    }
}
