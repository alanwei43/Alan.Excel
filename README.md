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

### Use 2

有时候给模型的属性添加注解(ExcelDesc)不是那么方便, 或者你需要允许用户自行配置映射关系, 比如第一个Excel其中一个头叫"分部", 后来改成了"店址", 如果使用注解就需要重新编译代码, 现在你可以使用下面的代码来注入你的映射关系
	
	import.InjectPropertyMap(new ExcelPropertyMap("StoreCity", "店址", typeof(string)));

而且你可以将多个Excel头映射到同一个属性:

	import.InjectPropertyMap(new ExcelPropertyMap("StoreCity", "店址", typeof(string)));
	import.InjectPropertyMap(new ExcelPropertyMap("StoreCity", "分部", typeof(string)));

现在Excel里的"店址"和"分部"都会映射到StoreCity.

### Update 1.2.5
在 1.2.5版本中新增以Fluent方式获取ExcelPropertyMap:
	
	var maps = Alan.Excel.Import.ExcelPropertyMap
		.Push("OverdueId", "逾期", typeof (string))
		.Push("LoanDate", "日期", typeof (DateTime))
		.Push("StoreAddress", "分部", typeof (string))
		.Push("Store", "门店", typeof (string))
		.Push("ApplyNo", "申请编号", typeof (string))
		.Push("Age", "年龄", typeof (int))
		.Get();

