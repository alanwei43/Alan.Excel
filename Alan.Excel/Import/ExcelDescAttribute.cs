using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Alan.Excel.Import
{
    public class ExcelDescAttribute : Attribute
    {
        public string Name { get; set; }
    }
}
