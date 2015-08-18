using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace Alan.Excel.Import
{
    public class ExcelImport<TModel>
        where TModel : class, new()
    {

        /// <summary>
        /// 记录异常信息
        /// </summary>
        private List<Exception> _exceptions;
        private void AddException(Exception ex)
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


        private Dictionary<string, Func<ExcelWorksheet, int, int, object>> _converts;

        /// <summary>
        /// 类型转换
        /// </summary>
        private Func<object, Type, object> _convertType = (cellValue, targetType) =>
        {
            return Convert.ChangeType(cellValue, targetType);
        };

        public ExcelImport() { }

        public ExcelImport(Dictionary<string, Func<ExcelWorksheet, int, int, object>> converts)
        {
            this._converts = converts;
        }

        /// <summary>
        /// 注入转换器
        /// </summary>
        /// <param name="typeFullName">匹配的类型名字全称</param>
        /// <param name="convert">转换委托(参数一次是: 当前的sheet, Row Index, Column Index)</param>
        public void InjectConvert(string typeFullName, Func<ExcelWorksheet, int, int, object> convert)
        {
            if (this._converts.ContainsKey(typeFullName)) return;
            this._converts[typeFullName] = convert;
        }

        /// <summary>
        /// 注入将 Excel Cell 里的值设置到Model时发生的类型转换
        /// </summary>
        /// <param name="convert">object:是Excel Cell值, Type: 目标属性的类型</param>
        public void InjectConvertType(Func<object, Type, object> convert)
        {
            this._convertType = convert;
        }

        /// <summary>
        /// 根据 Excel表头名字获取 此列数据的数据类型
        /// </summary>
        /// <param name="name">Excel表头名字</param>
        /// <returns></returns>
        private Type GetExcelType(string name)
        {
            Type retType = null;
            var model = new TModel();
            model.GetType().GetProperties().ToList().ForEach(property =>
          {
              var attribute = property.GetCustomAttributes(false).FirstOrDefault(att => att.GetType().FullName == typeof(ExcelDescAttribute).FullName);
              var desc = attribute as ExcelDescAttribute;
              if (desc == null) return;
              if (desc.Name == name) retType = property.PropertyType;
          });
            return retType ?? typeof(string);
        }

        /// <summary>
        /// 设置对象值
        /// </summary>
        /// <param name="model"></param>
        /// <param name="values"></param>
        private void SetModelValues(TModel model, Dictionary<string, object> values)
        {
            model.GetType().GetProperties().ToList().ForEach(property =>
            {
                var attribute = property.GetCustomAttributes(false).FirstOrDefault(att => att.GetType().FullName == typeof(ExcelDescAttribute).FullName);

                var desc = attribute as ExcelDescAttribute;
                if (desc == null) return;
                if (!values.ContainsKey(desc.Name)) return;

                var propType = property.PropertyType;
                var value = values[desc.Name];
                if (value == null) return;

                try
                {
                    var propertyValue = this._convertType(value, propType);
                    property.SetValue(model, propertyValue, null);
                }
                catch (Exception ex)
                {
                    this.AddException(ex);
                }

            });
        }

        /// <summary>
        /// 内置的转换器
        /// </summary>
        private Dictionary<string, Func<ExcelWorksheet, int, int, object>> GlobalConverts
        {
            get
            {
                var converts = new Dictionary<string, Func<ExcelWorksheet, int, int, object>>();
                converts.Add(typeof(DateTime).FullName, (sheet, row, column) => sheet.GetValue<DateTime>(row, column));
                return converts;
            }
        }

        /// <summary>
        /// 转换成Model列表
        /// </summary>
        /// <param name="sheet">ExcelWorksheet</param>
        /// <returns></returns>
        public List<TModel> ToModels(ExcelWorksheet sheet)
        {
            var models = new List<TModel>();
            for (var rowIndex = 2; rowIndex <= sheet.Dimension.Rows; rowIndex++)
            {
                var row = new Dictionary<string, object>();
                TModel model = new TModel();
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
                this.SetModelValues(model, row);
                models.Add(model);
            }
            return models;
        }


        /// <summary>
        /// 将某个Sheet转换成Models
        /// </summary>
        /// <param name="fileFullPath">Excel文件绝对路径</param>
        /// <param name="sheetName">Sheet名字</param>
        /// <returns></returns>
        public List<TModel> ToModels(string fileFullPath, string sheetName)
        {
            var models = new List<TModel>();
            ImportUtils.Sheet(fileFullPath, sheetName, sheet =>
            {
                models = this.ToModels(sheet);
            });
            return models;
        }

        /// <summary>
        /// 将某个Sheet转换成Models
        /// </summary>
        /// <param name="fileFullPath">Excel文件绝对路径</param>
        /// <param name="index">Sheet索引</param>
        /// <returns></returns>
        public List<TModel> ToModels(string fileFullPath, int index)
        {
            var models = new List<TModel>();
            ImportUtils.Sheet(fileFullPath, index, sheet =>
            {
                models = this.ToModels(sheet);
            });
            return models;
        }
    }
}
