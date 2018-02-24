package upb;

import java.util.ArrayList;
import java.util.List;

import org.apache.log4j.LogManager;
import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import comm.DataLineInfo;
import comm.ReadExcelUtil;
import comm.TableInfo;

public class ReadExeclSheet {
	private static Logger logger = LogManager.getLogger("Execl");   
	// 读表清单  获得 表信息列表 
	
	public List<TableInfo> readTableList (Workbook  workbook, String sheetName ) {
		logger.info("读表清单  获得 表信息列表  开始 ");
		Sheet readsheet = workbook.getSheet(sheetName);
		
		List<TableInfo> tableList = new ArrayList<TableInfo>();
		tableList.clear();
		int baserow = -1 ;
		int baseCol = -1 ;
		int rowNum = readsheet.getLastRowNum();
		
//		int seqCol = 0;
		int modCol = 1;
		int tabEnCol = 2;
		int tabChCol = 3;
		int tabDesCol = 4;
		int isTabCol = 5;
		int verCol = 6;
		int alterTimeCol = 7;
		for (int i=0; i<=rowNum && baserow == -1; i++){
			logger.debug("当前读取行: " + i);
			Row row = readsheet.getRow(i);
			int colNum = row.getLastCellNum();
			for(int j=0; j<colNum ; j++ ){
//				Cell cell = row.getCell(j);
				String content = ReadExcelUtil.getCellString(row, j);
				switch (content) {
				case "序号":
//					seqCol = j;
					baserow = i;
					baseCol = j;
					break;
				case "模块":
					modCol = j;
					break;
				case "表英文名称":
					tabEnCol = j;
					break;
				case "表中文名称":
					tabChCol = j;
					break;
				case "表功能描述":
					tabDesCol = j;
					break;
				case "表/视图":
					isTabCol = j;
					break;
				case "版本":
					verCol = j;
					break;
				case "变更时间":
					alterTimeCol = j;
					break;
				default:
					break;
				}
			}
		}
		logger.debug("获取到的 baserow :" + baserow);
		logger.debug("获取到的 baseCol :" + baseCol);
		
		if(baserow == rowNum ){
			return null;
		}
		
		logger.debug("开始读取TableList数据");
		for (int i=baserow+1; i<=rowNum; i++){

			TableInfo tbInfo = new TableInfo();
			Row row = readsheet.getRow(i);
			if(row == null ){
				continue;
			}
			String tabEnName = ReadExcelUtil.getCellString(row, tabEnCol);
			String tabChName = ReadExcelUtil.getCellString(row, tabChCol);
			if("".equals(tabEnName) && "".equals(tabChName)) {
				logger.warn("当前处理行为:" + i + "表名为空, 跳过本行" );
				continue;
			}
			logger.debug("当前处理行为:" + i);
//			logger.debug("当前处理行为:" + row.getCell(seqCol) );
			tbInfo.setSeqNo( i - baserow );
			tbInfo.setModel(ReadExcelUtil.getCellString(row, modCol) );
			tbInfo.setTabEnName(ReadExcelUtil.getCellString(row, tabEnCol).toUpperCase() );
			tbInfo.setTabChName(ReadExcelUtil.getCellString(row, tabChCol) );
			tbInfo.setTabDesc(ReadExcelUtil.getCellString(row, tabDesCol));
			tbInfo.setTabOrView(ReadExcelUtil.getCellString(row, isTabCol) );
			tbInfo.setVersion(ReadExcelUtil.getCellString(row, verCol));
			tbInfo.setAltTime(ReadExcelUtil.getCellString(row, alterTimeCol));
			//  20180125  读取超链接  
			if(row != null){
				Cell cell = row.getCell(tabEnCol);  
				if (cell != null ){
					String linkSheetName = "";
					//org.apache.poi.common.usermodel.Hyperlink link = cell.getHyperlink();
					Hyperlink link = cell.getHyperlink();
					
					if(link != null)
					{
						logger.info("AAAAAAAAAAA" + link.getAddress()) ;
						
						String tmpStr[] = link.getAddress().split("!");
						if(tmpStr.length > 0){
							linkSheetName = tmpStr[0];
							linkSheetName = linkSheetName.replaceAll("^'", "");
							linkSheetName = linkSheetName.replaceAll("'$", "");
						}else{
							linkSheetName = link.getAddress();
						}
						logger.info(linkSheetName) ;
						tbInfo.setLink(linkSheetName);  // 记录链接所在的Sheet名称
					}
				}
			}
			tableList.add(tbInfo);
		}
		
		logger.info("完成读取TableList数据");
		logger.info("读表清单  获得 表信息列表  完成 ");
		return tableList;
		
	}
	
	
	// 读表结构详细信息   
	public List<DataLineInfo> readTableDef (Workbook  workbook, String InSheetName , String tabEnName, String tabChName ) {
logger.info("读 " + tabEnName + " : " + tabChName + "表结构详细信息  开始 ");
		
		String sheetName = InSheetName  ;
		logger.info("英文表名 ["  + tabEnName + "] 中文名 [" + tabChName + "] Sheet 名[" + InSheetName + "]");
		Sheet readsheet = workbook.getSheet(sheetName);
		
		String  newTabChName =  null; 
		if (readsheet == null){
			if(tabChName.contains("/") || tabChName.contains("\\")){
				newTabChName = tabChName.replaceAll("/", "");
			}else {
				newTabChName = tabChName;
			}
			logger.info( "使用SheetName[" +  InSheetName + "]未找到, 使用英文表名 " + tabEnName + " 查找");
			
			sheetName = tabEnName  ;
			readsheet = workbook.getSheet(sheetName);
		}
		//通过英文名查找不到 , 使用中文名查找 
		if (readsheet == null){
			sheetName = newTabChName  ;
			logger.info("使用英文表名 "  + tabEnName + " 未查到 sheet, 使用中文名 " + sheetName + " 查找");
			readsheet = workbook.getSheet(sheetName);
		}
		//如果 中文名也查找不到, MPB系统 使用  newTabChName + "(" + tabEnName + ")"  格式
		if (readsheet == null){
			sheetName = newTabChName + "(" + tabEnName + ")"  ;
			logger.info("使用中文表名 "  + tabChName + " 未查到 sheet, 使用 " + sheetName + " 查找");
			readsheet = workbook.getSheet(sheetName);
			if(readsheet == null){
				sheetName = newTabChName + "（" + tabEnName + "）"  ;
				readsheet = workbook.getSheet(sheetName);
			}
			if(readsheet == null){
				sheetName = newTabChName + "（" + tabEnName + ")"  ;
				readsheet = workbook.getSheet(sheetName);
			}
			if(readsheet == null){
				// 中国人寿投保单信息表（mbpzrinfo)
				sheetName = newTabChName + "(" + tabEnName + "）"  ;
				readsheet = workbook.getSheet(sheetName);
			}
			
			else{
				logger.warn("读 " + tabEnName + "表,查找 " + sheetName  + "sheet成功");
			}
		}
		if(readsheet == null){
			logger.error("表 "+ tabEnName + " " + sheetName + " 未找到, 跳过 ");
			return null;
		}
		
		
		List<DataLineInfo> dataLineList = new ArrayList<DataLineInfo>();
		dataLineList.clear();
		int baserow = -1 ;
		int baseCol = -1 ;
		int rowNum = readsheet.getLastRowNum();
		
//		int seqNoCol   = 0;    
		int enNameCol  = 0;    
		int chNameCol  = 1;    
		int typeCol    = 2;    
		int isNullCol  = 3;    
		int isPKCol    = 4;    
		int isFKCol    = 100;    
		int defaultValueCol = 6;   
		int descCol = 7;           
		int dataStandardNoCol = 8 ; 
		int dataStandardNameCol = 9;
//		int odsTypeCol        
		for (int i=0; i<=rowNum && baserow == -1; i++){
			logger.debug("当前读取行: " + i);
			Row row = readsheet.getRow(i);
			int colNum = row.getLastCellNum();
			for(int j=0; j<colNum ; j++ ){
//				Cell cell = row.getCell(j);
				String content = ReadExcelUtil.getCellString(row, j);
				if(content.contains("字段英文名") ){
					baserow = i;
					baseCol = j;
					enNameCol = j;
				}else if (content.contains("字段英文名") || "字段英文名称".equals(content)){
					enNameCol = j;
				}else if (content.contains("字段中文名") || "字段中文名称".equals(content)){
					chNameCol = j;
				}else if (content.contains("字段类型") || "类型".equals(content) ){
					typeCol = j;
				}else if (content.contains("是否允许为空") || "可为空".equals(content)){
					isNullCol = j;
				}else if (content.contains("主键")){
					isPKCol = j;
				}else if (content.contains("外键")){
					isFKCol = j;
				}else if (content.contains("缺省值")  || "默认值".equals(content)){
					defaultValueCol = j;
				}else if (content.contains("字段说明") || content.contains("字段描述")){
					descCol = j;
				}else if (content.contains("数据标准编号") || "标准编号".equals(content)){
					dataStandardNoCol = j;
				}else if (content.contains("数据标准中文名称") || "标准中文名称".equals(content)){
					dataStandardNameCol = j;
				}
			}
		}
		logger.debug("获取到的 baserow :" + baserow);
		logger.debug("获取到的 baseCol :" + baseCol);
		
		if(baserow == rowNum ){
			return null;
		}
		
		DataLineInfo dataLine = null;
		for (int i=baserow+1; i<=rowNum; i++){
			Row row = readsheet.getRow(i);
			if(row == null){
				continue;
			}
			else if (
					row.getCell(enNameCol).getStringCellValue().contains("索引名称") 
					|| row.getCell(enNameCol).getStringCellValue().contains("索引信息") 
					|| "索引信息".equals(row.getCell(enNameCol).getStringCellValue())
					|| "索引名".equals(row.getCell(enNameCol).getStringCellValue())) {
				break;
			}
			logger.debug("row.getCell(enNameCol).getStringCellValue()" + row.getCell(enNameCol).getStringCellValue());
			String columnEnName =  ReadExcelUtil.getCellString(row, enNameCol);
			String columnChName =  ReadExcelUtil.getCellString(row, chNameCol);
			
			if("".equals(columnEnName) && "".equals(columnChName)){
				dataLine.setDesc( dataLine.getDesc() + "\n" + ReadExcelUtil.getCellString(row, descCol) );
			}
			else{
				dataLine = new DataLineInfo();

//				logger.debug(tabEnName + " 当前处理行为:" + i + " :" + row.getOutlineLevel());
				dataLine.setSeqNo(i - baserow );              
				dataLine.setEnName          (columnEnName.toUpperCase());
				dataLine.setChName          (columnChName);
				dataLine.setType            (ReadExcelUtil.getCellString(row, typeCol  ).toUpperCase());
				dataLine.setIsNull          (ReadExcelUtil.getCellString(row, isNullCol));
				dataLine.setIsPK            (ReadExcelUtil.getCellString(row, isPKCol  ));
				dataLine.setIsFK            (ReadExcelUtil.getCellString(row, isFKCol  ));
				dataLine.setDefaultValue    (ReadExcelUtil.getCellString(row, defaultValueCol)); 
				dataLine.setDesc            (ReadExcelUtil.getCellString(row, descCol)); 
				dataLine.setDataStandardNo  (ReadExcelUtil.getCellString(row, dataStandardNoCol)); 
				dataLine.setDataStandardName(ReadExcelUtil.getCellString(row, dataStandardNameCol)); 
				dataLineList.add(dataLine);
			}
		}
		
		logger.info("读 " + tabEnName + "表结构详细信息  完成 ");
		return dataLineList;
	}
	
	

}
