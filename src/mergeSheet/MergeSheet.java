package mergeSheet;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.log4j.LogManager;
import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import comm.CreateExeclUtil;
import comm.ReadExcelUtil;

/**
 * 合并多个文件中的多个 Sheet 到 同一个Sheet 中 
 * @author ding_kaiye
 *
 */
public class MergeSheet {

	private static Logger logger = LogManager.getLogger("Execl"); 
	
	public Sheet getSheet(String fileName, String sheetName) throws Exception{
		Sheet  sheet = null; 
		Workbook workbook = CreateExeclUtil.getReadWorkbook(fileName);
		if (workbook == null){
			logger.error("打开文件错误[" + fileName + "]");
			throw new Exception("打开文件错误[" + fileName + "]");
		}
		sheet = workbook.getSheet(sheetName);
		if (sheet == null){
			logger.error("打开Sheet错误[" + fileName + "] [" + sheetName + "]");
			return null;
		}
		return  sheet ; 
	}

	private Sheet  AddSheeRows (Sheet targetSheet, Sheet sourceSheet){
		
		if(targetSheet == null){
			logger.error("targetSheet is null");
			return targetSheet ;
		}
		if(sourceSheet == null){
			logger.error("sourceSheet is null");
			return targetSheet ;
		}
		
		int rowCnt = 0;
		int soureRowNum = sourceSheet.getLastRowNum();
		int targetRowNum = targetSheet.getLastRowNum();
		Row sourceRow = null;
		Row targetRow = null;
		int colNum = 0;
	//获取sheet , 
		for (int i=0; i<=soureRowNum;i++){
			sourceRow = sourceSheet.getRow(i);
			//logger.debug("读取第 "+ i + "行");
			
			if(sourceRow != null){
				//检查所有字段是否为空
				colNum = sourceRow.getLastCellNum();
				boolean recvflg = false;
				for(int j=0; j<colNum ; j++ ){
					String content = ReadExcelUtil.getCellString(sourceRow, j);
					if(content != null && !"".equals(content.trim()) ){
						recvflg = true;  
						break;
					}
				}
				if(recvflg == true){
					targetRow = targetSheet.createRow(targetRowNum + rowCnt );
					rowCnt ++;
					//targetRow = targetSheet.createRow(1);
					logger.debug("第 "+ i + "行追加到Sheet 第 " + (targetRowNum + rowCnt) );
					for(int k=0; k<colNum ; k++ ){
						String content = ReadExcelUtil.getCellString(sourceRow, k);
						targetRow.createCell(k).setCellValue(content);
						logger.trace("第 "+ targetRowNum + rowCnt + "行, 第" + k + "列添加" + "["  + content + "]");
					}
				}
			}
		}
	// 读取当前sheet 到 新sheet
		return targetSheet ;
	}
	
	
	public void addSheetConnect(String targetFile, String targetSheetName, 
			String sourceFile, String sourceSheetName ) throws Exception{
		logger.info("将" + sourceFile + "[" + sourceSheetName + "] 写入 " + targetFile + "[" + targetSheetName + "]  开始 ");
		
		Workbook  targetWorkbook = CreateExeclUtil.getReadWorkbook(targetFile);
//		Workbook  sourceWorkbook = CreateExeclUtil.getReadWorkbook(sourceFile);
		Sheet targetSheet = getSheet(targetFile, targetSheetName);
		if(targetSheet == null){
			logger.error("打开Sheet错误[" + targetSheetName + "]");
			targetSheet = targetWorkbook.createSheet(targetSheetName);
		}
		Sheet sourceSheet = getSheet(sourceFile, sourceSheetName);
		//targetSheet = AddSheeRows(targetSheet, sourceSheet, targetFile);
		targetSheet = AddSheeRows(targetSheet, sourceSheet);
		
		writeWorkbookToFile(targetSheet.getWorkbook(),  targetFile  );
		logger.info("将" + sourceFile + "[" + sourceSheetName + "] 写入 " + targetFile + "[" + targetSheetName + "]  结束 ");
		
	}
	
	
	
	public static boolean writeWorkbookToFile(Workbook workbook, String outFile) {
		logger.info("将 workbook 写入文件 " + outFile  + "  开始 ");
		FileOutputStream fOut =  null;
		try {
			File file = new File(outFile);
//			if (file.exists() && file.isFile()) {
//				logger.info(outFile + "已存在, 开始删除文件 ");
//				file.delete();
//				logger.info(outFile + "已存在, 删除文件完成 ");
//			}
			fOut = new FileOutputStream(outFile);
			workbook.write(fOut);
			fOut.flush();  
			fOut.close();  // 操作结束，关闭文件  
		} catch (FileNotFoundException e) {
			logger.error("FileNotFoundException", e);
			return false;
		} catch (IOException e) {
			logger.error("IOException", e);
			return false;
		}
		logger.info("将 workbook 写入文件 " + outFile  + "  完成");
		return true;

	}
	
	

}




