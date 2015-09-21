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
            var fileFullPath = @"E:\Shared\IAP.xlsx";
            FileStream fs = new FileStream(fileFullPath, FileMode.Open, FileAccess.Read, FileShare.Read);
            var import = new ExcelImportModel<Ysd.DataAccessLayer.Models.Biz_Wealth>();
            import.InjectPropertyMaps(GetExcelPropertyMap());
            var models = import.ToModels(fs, 1);
        }



        public static List<Alan.Excel.Import.ExcelPropertyMap> GetExcelPropertyMap()
        {
            var maps = Alan.Excel.Import.ExcelPropertyMap
              .Push("Subsectors", "城市", typeof(string))
              .Push("Stores", "门店", typeof(string))
              .Push("WealthDate", "日期", typeof(DateTime))
              .Push("CustomerNo", "客户编号", typeof(string))
              .Push("CustomerName", "客户姓名", typeof(string))
              .Push("Sex", "性别", typeof(string))
              .Push("ContractNo", "合同编号", typeof(string))
              .Push("LendNo", "出借编号", typeof(string))
              .Push("ProductName", "产品", typeof(string))
              .Push("CustomerAttribute", "客户性质", typeof(string))
              .Push("FundAmount", "资金额度", typeof(decimal))
              .Push("WealthMonths", "投资年限", typeof(decimal))
              .Push("Yield", "年化收益", typeof(decimal))
              .Push("IncomeDate", "到账时间", typeof(DateTime))
              .Push("InterestDate", "起息时间", typeof(DateTime))
              .Push("ClosedDate", "封闭期到期日期", typeof(DateTime))
              .Push("BillDay", "每月出账单日", typeof(string))
              .Push("RedeemInfo", "赎回情况", typeof(string))
              .Push("Tel", "联系方式", typeof(string))
              .Push("CardNo", "身份证号码", typeof(string))
              .Push("MailWay", "是否邮寄", typeof(string))
              .Push("MailAddress", "纸质邮箱", typeof(string))
              .Push("Email", "电子邮箱", typeof(string))
              .Push("BankName", "所属银行", typeof(string))
              .Push("BankBranch", "所属支行", typeof(string))
              .Push("BankNo", "银行卡号", typeof(string))
              .Push("IfContinue", "是否续投", typeof(string))
              .Push("Consultant", "理财顾问", typeof(string))
              .Push("Manager", "团队经理", typeof(string))
              .Push("PromotionChannel", "推广渠道", typeof(string))
              .Push("ICEContract", "紧急联系人", typeof(string))
              .Push("ICEPhone", "紧急电话", typeof(string))
              .Push("Remarks", "备注", typeof(string))
              .Push("AccountNo", "银行账户号", typeof(string))
              .Get();

            return maps;
        }

    }


}
