package mbp;

import java.io.File;
import java.util.List;

import org.apache.log4j.LogManager;
import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import comm.DataLineInfo;
import comm.ReadExcelUtil;
import comm.CreateExeclUtil;
import comm.TableInfo;

public class CreateMBPExecl {
	private static Logger logger = LogManager.getLogger("Execl");   
	public static void main(String[] argv){
		if(argv.length != 4){
			logger.error("参数数量不对\n");
			logger.error("源系统数据字典文件   模板文件    输出文件  模块名称");
			System.exit(-1);
		}
		String fileName = argv[0];  
		String mufileName = argv[1];  
		String outFile = argv[2];   
		String muName = argv[3];
	
		boolean errorFlg = false;
		if(fileName == null || "".equals(fileName)){
			logger.error("数据字典文件不能为空\n");
			errorFlg = true;
		}
		if(mufileName == null || "".equals(mufileName)){
			logger.error("模板文件不能为空\n");
			errorFlg = true;
		}
		if(outFile == null || "".equals(outFile)){
			logger.error("输出文件 不能为空\n");
			errorFlg = true;
		}
		
		File file=new File(fileName);
        if(!file.exists() || !file.isFile()){
        	logger.error("数据字典文件: " + fileName + " 不存在");
        	errorFlg = true;
		}
        file=new File(mufileName);
        if(!file.exists() || !file.isFile()){
        	logger.error("模板文件        : " + mufileName + " 不存在");
        	errorFlg = true;
		}

		if(errorFlg){
			logger.error("   请检查!!!");
			System.exit(-1);
		}
		
		logger.info("数据字典文件:" + fileName);
		logger.info("模板文件        :" + mufileName);
		logger.info("输出文件         :" + outFile);
		
		ReadExeclSheet worker = new ReadExeclSheet();
		
		Workbook  workbook = CreateExeclUtil.getReadWorkbook(fileName);
		List<TableInfo> tablelist = null;
		try{
			tablelist = worker.readTableList(workbook, "表（视图)清单");
		} catch (Exception e){
			logger.error("读取表清单出错" , e);
			throw e;
		}
		
		ReadExeclSheet readExeclSheet = new ReadExeclSheet ();
		for(TableInfo tableInfoCur : tablelist) {
			String tabEnName = tableInfoCur.getTabEnName();
			String tabChName = tableInfoCur.getTabChName();
			
			String SheetName = tableInfoCur.getLink();
			
			List<DataLineInfo> dataList = readExeclSheet.readTableDef ( workbook, SheetName, tabEnName,  tabChName ) ; 
			tableInfoCur.setDatalineList(dataList);
		}
		ReadExcelUtil.CreatTableListOdsType(tablelist);
		//String[] splits = {"\n",     ";",  "；",      "," , "，" ,     ":", "：",    "-", "–",  " "};   //枚举值得分隔符 
		String[] splits = {"\n",     ";",  "；",  "-", "–",   "," , "，" ,   ":", "：",  };   //枚举值得分隔符 
		ReadExcelUtil.CreatTableListEnum(tablelist, splits, splits);
		
		Workbook  resultWorkbook = CreateExeclUtil.getReadWorkbook(mufileName);
		Sheet emenuSheet = resultWorkbook.getSheet("源系统代码");
		Sheet tableListSheet = resultWorkbook.getSheet("源表");
		Sheet ableDefSheet = resultWorkbook.getSheet("源字段");
		
		CreateExeclUtil.CreatTableDefSheet("UPB", ableDefSheet, tablelist);
		CreateExeclUtil.CreatTableListSheet("UPB", muName, tableListSheet, tablelist );
		CreateExeclUtil.CreatEmenuSheet("UPB", emenuSheet, tablelist);
		
		CreateExeclUtil.writeWorkbookToFile(resultWorkbook, outFile);
		
		logger.debug(tablelist);
	}

}
