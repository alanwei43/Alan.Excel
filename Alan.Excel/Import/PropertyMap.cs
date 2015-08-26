using System;
using System.Collections.Generic;

namespace Alan.Excel.Import
{
    /// <summary>
    /// Model属性和Excel Header Name之间的映射
    /// </summary>
    public class ExcelPropertyMap
    {
        public ExcelPropertyMap() { }

        /// <summary>
        /// 实例化 Map
        /// </summary>
        /// <param name="propName">Model Property Name</param>
        /// <param name="headerName">Excel Header Name</param>
        /// <param name="propType">Model Property Type</param>
        public ExcelPropertyMap(string propName, string headerName, Type propType)
        {
            this.ModelPropertyName = propName;
            this.ExcelHeaderName = headerName;
            this.PropertyType = propType;
        }

        /// <summary>
        /// Model属性
        /// </summary>
        public string ModelPropertyName { get; set; }

        /// <summary>
        /// Excel header name
        /// </summary>
        public string ExcelHeaderName { get; set; }

        /// <summary>
        /// 属性类型
        /// 如果为空 则使用反射时获取的Model的属性的类型
        /// </summary>
        public Type PropertyType { get; set; }

        public static ExcelPropertyMapHelper Push(string propName, string headerName, Type propType)
        {
            return (new ExcelPropertyMapHelper()).Push(propName, headerName, propType);
        }
        public static ExcelPropertyMapHelper Push(ExcelPropertyMap map)
        {
            return (new ExcelPropertyMapHelper()).Push(map);
        }
    }

    public class ExcelPropertyMapHelper
    {
        /// <summary>
        /// ExcelPropertyMaps
        /// </summary>
        public readonly List<ExcelPropertyMap> _maps;

        public ExcelPropertyMapHelper()
        {
            this._maps = new List<ExcelPropertyMap>();
        }

        /// <summary>
        /// 添加
        /// </summary>
        /// <param name="propName">Model Property Name</param>
        /// <param name="headerName">Excel Header Name</param>
        /// <param name="propType">Model Property Type</param>
        /// <returns></returns>
        public ExcelPropertyMapHelper Push(string propName, string headerName, Type propType)
        {

            return this.Push(new ExcelPropertyMap(propName, headerName, propType));
        }

        /// <summary>
        /// 添加
        /// </summary>
        /// <param name="map"></param>
        /// <returns></returns>
        public ExcelPropertyMapHelper Push(ExcelPropertyMap map)
        {
            this._maps.Add(map);
            return this;
        }

        public List<ExcelPropertyMap> Get()
        {
            return this._maps;
        }

    }
}