using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;

namespace Alan.Excel.Import
{
    /// <summary>
    /// Excel导入的基类, 可以通过继承这个类来实现自定义的功能
    /// ExcelImportModel就是这个类的一个实现
    /// </summary>
    public class ExcelImport
    {
        #region Exception Record
        /// <summary>
        /// 记录异常信息
        /// </summary>
        private List<Exception> _exceptions;
        protected void AddException(Exception ex)
        {
            if (this._exceptions == null) this._exceptions = new List<Exception>();
            this._exceptions.Add(ex);
        }
        /// <summary>
        /// 获取异常信息
        /// </summary>
        /// <returns></returns>
        public List<Exception> GetExceptions()
        {
            return this._exceptions ?? new List<Exception>();
        }
        #endregion



        /// <summary>
        /// 获取Sheet单元格的值
        /// </summary>
        private Dictionary<string, Func<ExcelWorksheet, int, int, object>> _converts;

        /// <summary>
        /// Excel 头名字 与 Model 属性名字之间的映射关系
        /// </summary>
        protected List<ExcelPropertyMap> PropertyMaps;

        protected ExcelImport()
        {
            this.PropertyMaps = new List<ExcelPropertyMap>();
            this._exceptions = new List<Exception>();
            this._converts = new Dictionary<string, Func<ExcelWorksheet, int, int, object>>();
        }

        /// <summary>
        /// 实例化 属性映射
        /// </summary>
        /// <param name="propMaps">属性映射</param>
        public ExcelImport(List<ExcelPropertyMap> propMaps)
        {
            this.PropertyMaps = propMaps;
        }

        /// <summary>
        /// 实例化 属性映射, 转换器
        /// </summary>
        /// <param name="propMaps">属性映射</param>
        /// <param name="converts">转换器</param>
        public ExcelImport(
            List<ExcelPropertyMap> propMaps,
            Dictionary<string, Func<ExcelWorksheet, int, int, object>> converts) : this(propMaps)
        {
            this._converts = converts;
        }

        #region 注入 修改默认实现

        public void ReplaceGetCellValue(Dictionary<string, Func<ExcelWorksheet, int, int, object>> convert)
        {
            this._converts = convert;

        }

        /// <summary>
        /// 注入转换器
        /// </summary>
        /// <param name="typeFullName">匹配的类型名字全称</param>
        /// <param name="convert">转换委托(参数一次是: 当前的sheet, Row Index, Column Index)</param>
        public bool InjectGetCellValue(string typeFullName, Func<ExcelWorksheet, int, int, object> convert)
        {
            if (this._converts == null) this._converts = new Dictionary<string, Func<ExcelWorksheet, int, int, object>>();

            if (this._converts.ContainsKey(typeFullName)) return false;
            this._converts[typeFullName] = convert;
            return true;
        }

        /// <summary>
        /// 注入 自己的映射
        /// </summary>
        /// <param name="maps"></param>
        public void ReplacePropertyMap(List<ExcelPropertyMap> maps)
        {
            this.PropertyMaps = maps ?? new List<ExcelPropertyMap>();
        }

        /// <summary>
        /// 注入自己的映射
        /// </summary>
        /// <param name="map">映射关系</param>
        public void InjectPropertyMap(ExcelPropertyMap map)
        {
            if (this.PropertyMaps == null) this.PropertyMaps = new List<ExcelPropertyMap>();
            this.PropertyMaps.Add(map);
        }

        /// <summary>
        /// 一次注入多个自己的映射
        /// </summary>
        /// <param name="maps"></param>
        public void InjectPropertyMaps(List<ExcelPropertyMap> maps)
        {
            if (this.PropertyMaps == null) this.PropertyMaps = new List<ExcelPropertyMap>();
            this.PropertyMaps.AddRange(maps);
        }
        #endregion

        /// <summary>
        /// 根据 Excel表头名字获取 此列数据的数据类型
        /// </summary>
        /// <param name="name">Excel表头名字</param>
        /// <returns></returns>
        protected Type GetExcelType(string name)
        {
            var propertyMap = this.PropertyMaps.FirstOrDefault(propMap => propMap.ExcelHeaderName == name);
            if (propertyMap == null) return typeof(string);
            return propertyMap.PropertyType;
        }




        /// <summary>
        /// 获取所有的行
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        public List<Dictionary<string, object>> GetRows(ExcelWorksheet sheet)
        {
            var rows = new List<Dictionary<string, object>>();
            for (var rowIndex = 2; rowIndex <= sheet.Dimension.Rows; rowIndex++)
            {
                var row = new Dictionary<string, object>();

                for (var columnIndex = 1; columnIndex <= sheet.Dimension.Columns; columnIndex++)
                {
                    var cellName = sheet.GetValue(1, columnIndex);
                    if (cellName == null) continue;
                    var cellNameString = cellName.ToString();

                    object cellValue = sheet.GetValue<string>(rowIndex, columnIndex);
                    var cellType = this.GetExcelType(cellNameString);

                    if (this._converts != null && this._converts.ContainsKey(cellType.FullName))
                    {
                        //优先使用用户定义的转换器
                        cellValue = this._converts[cellType.FullName](sheet, rowIndex, columnIndex);
                    }
                    else
                    {
                        if (this.GlobalConverts.ContainsKey(cellType.FullName))
                        {
                            //如果没有匹配的 使用内置的转换器
                            cellValue = this.GlobalConverts[cellType.FullName](sheet, rowIndex, columnIndex);
                        }
                    }


                    if (row.ContainsKey(cellNameString))
                    {
                        row.Add(cellNameString + "1", cellValue);
                        continue;
                    }

                    row.Add(cellNameString, cellValue);
                }
                rows.Add(row);
            }
            return rows;
        }


        /// <summary>
        /// 内置的转换器
        /// </summary>
        private Dictionary<string, Func<ExcelWorksheet, int, int, object>> GlobalConverts
        {
            get
            {
                var converts = new Dictionary<string, Func<ExcelWorksheet, int, int, object>>();
                converts.Add(typeof(DateTime).FullName, (sheet, row, column) =>
                {
                    var value = sheet.GetValue<DateTime>(row, column);
                    if (value == default(DateTime))
                        return new DateTime(1970, 1, 1);
                    return value;
                });
                return converts;
            }
        }
    }
}
