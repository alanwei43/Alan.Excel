using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Alan.Excel.Import;

namespace Alan.Excel.Test
{
    class Program
    {
        static void Main(string[] args)
        {
            var fileFullPath = @"E:\Shared\1.xlsx";
            FileStream fs = new FileStream(fileFullPath, FileMode.Open, FileAccess.Read, FileShare.Read);
            var import = new ExcelImportModel<Model>();
            var models = import.ToModels(fs, 1);
        }

        public class Model
        {
            [ExcelDesc(Name = "城市")]
            public string City { get; set; }
            [ExcelDesc(Name = "门店")]
            public string Door { get; set; }
            [ExcelDesc(Name = "日期")]
            public DateTime Date { get; set; }
            [ExcelDesc(Name = "客户编号")]
            public string CustomerId { get; set; }

            [ExcelDesc(Name = "客户编号")]
            public string CustomerName { get; set; }
        }
    }


}
