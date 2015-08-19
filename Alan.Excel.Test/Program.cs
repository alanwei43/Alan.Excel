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
            var fileFullPath = @"E:\Projects\AspNetMvc\AspNetMvc\Content\2015year.xlsx";


            var import = new Alan.Excel.Import.ExcelImportModel<RepaymentModel>();
            var models = import.ToModels(fileFullPath, "201506借款客户总表");
        }
    }


    public class RepaymentModel
    {
        [ExcelDesc(Name = "逾期")]
        public string Overlay { get; set; }

        [ExcelDesc(Name = "日期")]
        public DateTime Date { get; set; }

        [ExcelDesc(Name = "分部")]
        public string StoreCity { get; set; }

        [ExcelDesc(Name = "门店")]
        public string StoreName { get; set; }

        [ExcelDesc(Name = "申请编号")]
        public string ApplyNo { get; set; }
        [ExcelDesc(Name = "借款人姓名")]
        public string LoanerName { get; set; }
        [ExcelDesc(Name = "身份证")]
        public string IdCardNo { get; set; }

        [ExcelDesc(Name = "客户邮箱")]
        public string CustomerEmail { get; set; }

        [ExcelDesc(Name = "手机")]
        public string PhoneNumber { get; set; }

        [ExcelDesc(Name = "期数（元）")]
        public int MonthCount { get; set; }


        [ExcelDesc(Name = "签约金额")]
        public decimal SignAmount { get; set; }

        [ExcelDesc(Name = "管理总费用")]
        public decimal ManageAmout { get; set; }

        [ExcelDesc(Name = "放款金额")]
        public decimal LoanAmount { get; set; }

        [ExcelDesc(Name = "费率")]
        public float Rate { get; set; }

        [ExcelDesc(Name = "月还款金额")]
        public decimal AmountPerMonth { get; set; }

        [ExcelDesc(Name = "签约日期")]
        public DateTime SignDate { get; set; }

        [ExcelDesc(Name = "审核人")]
        public string Checker { get; set; }
        [ExcelDesc(Name = "审核人1")]
        public string Checker1 { get; set; }

    }


}
