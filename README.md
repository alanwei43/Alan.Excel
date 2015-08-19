# Alan.Excel
Alan.Excel

### Nuget Address

  https://www.nuget.org/packages/Alan.Excel/
  
### Install
  
  Install-Package Alan.Excel
  
### Use

    //using Alan.Excel.Import
    var fileFullPath = Server.MapPath("~/Content/2015year.xlsx");
    var import = new Alan.Excel.Import.ExcelImportModel<RepaymentModel>();
    List<RepaymentModel> models = import.ToModels(fileFullPath, "sheetname");
  
Model定义如下:


    public class RepaymentModel
    {
      [ExcelDesc(Name = "逾期")]
      public string Overlay { get; set; }
      [ExcelDesc(Name = "日期")]
      public DateTime Date { get; set; }
      [ExcelDesc(Name = "分部")]
      public string StoreCity { get; set; }
    }
  
  
其中ExcelDesc注解里的Name是Excel里的头名称.
