using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace Alan.Excel.Import
{
    /// <summary>
    /// Excel导入生成Model
    /// </summary>
    /// <typeparam name="TModel">实体类型</typeparam>
    public class ExcelImportModel<TModel> : ExcelImport
        where TModel : class, new()
    {


        /// <summary>
        /// 类型转换
        /// </summary>
        private Func<object, Type, object> _convertType = (cellValue, targetType) =>
        {
            return Convert.ChangeType(cellValue, targetType);
        };


        public ExcelImportModel()
        {
            //初始化 _propertyMaps
            if (this.PropertyMaps == null) this.PropertyMaps = new List<ExcelPropertyMap>();
            var model = new TModel();
            model.GetType().GetProperties().ToList().ForEach(property =>
            {
                var attribute = property.GetCustomAttributes(false).FirstOrDefault(att => att.GetType().FullName == typeof(ExcelDescAttribute).FullName);

                var desc = attribute as ExcelDescAttribute;
                if (desc == null) return;

                this.PropertyMaps.Add(new ExcelPropertyMap()
                {
                    ExcelHeaderName = desc.Name,
                    ModelPropertyName = property.Name,
                    PropertyType = property.PropertyType
                });
            });
        }

        public ExcelImportModel(Dictionary<string, Func<ExcelWorksheet, int, int, object>> converts) : this()
        {
            this.ReplaceGetCellValue(converts);
        }

        /// <summary>
        /// 注入将 Excel Cell 里的值设置到Model时发生的类型转换
        /// </summary>
        /// <param name="convert">object:是Excel Cell值, Type: 目标属性的类型</param>
        public void InjectSetModelValue(Func<object, Type, object> convert)
        {
            this._convertType = convert;
        }


        /// <summary>
        /// 设置对象值
        /// </summary>
        /// <param name="dicts">属性:值</param>
        private TModel ToModel(Dictionary<string, object> dicts)
        {
            var model = new TModel();
            model.GetType().GetProperties().ToList().ForEach(property =>
            {
                var propertyMap = this.PropertyMaps.FirstOrDefault(propMap => propMap.ModelPropertyName == property.Name);
                if (propertyMap == null) return;    //在映射里找不到这个属性
                if (!dicts.ContainsKey(propertyMap.ExcelHeaderName)) return;   //字典里没有对应的Excel值
                var value = dicts[propertyMap.ExcelHeaderName];
                if (value == null) return;      //字典里得值为空

                try
                {
                    var propertyValue = this._convertType(value, propertyMap.PropertyType ?? property.PropertyType);
                    property.SetValue(model, propertyValue, null);
                }
                catch (Exception ex)
                {
                    this.AddException(ex);
                }
            });
            return model;
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
            var rows = this.GetRows(sheet);
            return rows.Select(this.ToModel).ToList();
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
