using System;

namespace Alan.Excel.Import
{
    /// <summary>
    /// Model属性和Excel Header Name之间的映射
    /// </summary>
    public class ExcelPropertyMap
    {
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
    }
}