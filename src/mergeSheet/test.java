package mergeSheet;

public class test {

	
	public static void main(String agrv[]) throws Exception{
		
		// 写入文件 
		String tergetfile = "D:\\Temp\\AML\\All.xls";
		
		MergeSheet ms = new MergeSheet();
		
		
	//	ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls", "源字段", "D:\\Temp\\AML\\ODS标准数据结构文档(缓冲层及系统记录层)-AGL.xls", "源字段");
		//ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls", "源字段", "D:\\Temp\\AML\\ODS标准数据结构文档(缓冲层及系统记录层)-ATM.xls", "源字段");
		ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls","源字段","D:\\Temp\\AML\\ODS标准数据结构文档(缓冲层及系统记录层)-ECIF.xls","源字段");
		ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls","源字段","D:\\Temp\\AML\\ODS标准数据结构文档(缓冲层及系统记录层)-POS.xls","源字段");
		ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls","源字段","D:\\Temp\\AML\\ODS标准数据结构文档(缓冲层及系统记录层)-TFS.xls","源字段");
		ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls","源字段","D:\\Temp\\AML\\ODS标准数据结构文档(缓冲层及系统记录层)-HR.xls","源字段");
		ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls","源字段","D:\\Temp\\AML\\ODS标准数据结构文档(缓冲层及系统记录层)-AGL.xls","源字段");
		ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls","源字段","D:\\Temp\\AML\\ODS标准数据结构文档(缓冲层及系统记录层)-UPB.xls","源字段");
		ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls","源字段","D:\\Temp\\AML\\ODS标准数据结构文档(缓冲层及系统记录层)-OBM.xls","源字段");
		ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls","源字段","D:\\Temp\\AML\\ODS标准数据结构文档(缓冲层及系统记录层)-TBC.xls","源字段");
		ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls","源字段","D:\\Temp\\AML\\ODS标准数据结构文档(缓冲层及系统记录层)-ECD.xls","源字段");
		ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls","源字段","D:\\Temp\\AML\\ODS标准数据结构文档(缓冲层及系统记录层)-CMS.xls","源字段");
		ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls","源字段","D:\\Temp\\AML\\ODS标准数据结构文档(缓冲层及系统记录层)-ATM.xls","源字段");
		
		ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls","源字段","D:\\Temp\\AML\\ODS标准数据结构文档(缓冲层及系统记录层)-CBS-不良资产分册.xls","源字段");
		ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls","源字段","D:\\Temp\\AML\\ODS标准数据结构文档(缓冲层及系统记录层)-CBS-个人结售汇&外汇买卖分册.xls","源字段");
		ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls","源字段","D:\\Temp\\AML\\ODS标准数据结构文档(缓冲层及系统记录层)-CBS-利率分册.xls","源字段");
		ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls","源字段","D:\\Temp\\AML\\ODS标准数据结构文档(缓冲层及系统记录层)-CBS-外围分册.xls","源字段");
		ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls","源字段","D:\\Temp\\AML\\ODS标准数据结构文档(缓冲层及系统记录层)-CBS-机构分册.xls","源字段");
		ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls","源字段","D:\\Temp\\AML\\ODS标准数据结构文档(缓冲层及系统记录层)-CBS-柜员分册.xls","源字段");
		ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls","源字段","D:\\Temp\\AML\\ODS标准数据结构文档(缓冲层及系统记录层)-CBS-现金分册.xls","源字段");
		ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls","源字段","D:\\Temp\\AML\\ODS标准数据结构文档(缓冲层及系统记录层)-CBS-结算分册.xls","源字段");
		ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls","源字段","D:\\Temp\\AML\\ODS标准数据结构文档(缓冲层及系统记录层)-CBS-股金分册.xls","源字段");
		ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls","源字段","D:\\Temp\\AML\\ODS标准数据结构文档(缓冲层及系统记录层)-CBS-费用分册.xls","源字段");
		ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls","源字段","D:\\Temp\\AML\\ODS标准数据结构文档(缓冲层及系统记录层)-CBS-资产证券化分册.xls","源字段");
		ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls","源字段","D:\\Temp\\AML\\ODS标准数据结构文档(缓冲层及系统记录层)-CBS-资金分册.xls","源字段");
		ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls","源字段","D:\\Temp\\AML\\ODS标准数据结构文档(缓冲层及系统记录层)-CBS_会计分册.xls","源字段");
		ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls","源字段","D:\\Temp\\AML\\ODS标准数据结构文档(缓冲层及系统记录层)-CBS_凭证分册.xls","源字段");
		ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls","源字段","D:\\Temp\\AML\\ODS标准数据结构文档(缓冲层及系统记录层)-CBS_存款分册.xls","源字段");
		ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls","源字段","D:\\Temp\\AML\\ODS标准数据结构文档(缓冲层及系统记录层)-CBS_贷款分册.xls","源字段");
		ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls","源字段","D:\\Temp\\AML\\ODS标准数据结构文档(缓冲层及系统记录层)-CBS_通用分册.xls","源字段");
		ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls","源字段","D:\\Temp\\AML\\ODS标准数据结构文档(缓冲层及系统记录层)-CBS_银行卡分册.xls","源字段");
	
	}
	
	
	
	
	
	
}
