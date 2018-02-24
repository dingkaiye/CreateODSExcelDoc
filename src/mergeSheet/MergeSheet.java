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
 * �ϲ�����ļ��еĶ�� Sheet �� ͬһ��Sheet �� 
 * @author ding_kaiye
 *
 */
public class MergeSheet {

	private static Logger logger = LogManager.getLogger("Execl"); 
	
	public Sheet getSheet(String fileName, String sheetName) throws Exception{
		Sheet  sheet = null; 
		Workbook workbook = CreateExeclUtil.getReadWorkbook(fileName);
		if (workbook == null){
			logger.error("���ļ�����[" + fileName + "]");
			throw new Exception("���ļ�����[" + fileName + "]");
		}
		sheet = workbook.getSheet(sheetName);
		if (sheet == null){
			logger.error("��Sheet����[" + fileName + "] [" + sheetName + "]");
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
	//��ȡsheet , 
		for (int i=0; i<=soureRowNum;i++){
			sourceRow = sourceSheet.getRow(i);
			//logger.debug("��ȡ�� "+ i + "��");
			
			if(sourceRow != null){
				//��������ֶ��Ƿ�Ϊ��
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
					logger.debug("�� "+ i + "��׷�ӵ�Sheet �� " + (targetRowNum + rowCnt) );
					for(int k=0; k<colNum ; k++ ){
						String content = ReadExcelUtil.getCellString(sourceRow, k);
						targetRow.createCell(k).setCellValue(content);
						logger.trace("�� "+ targetRowNum + rowCnt + "��, ��" + k + "�����" + "["  + content + "]");
					}
				}
			}
		}
	// ��ȡ��ǰsheet �� ��sheet
		return targetSheet ;
	}
	
	
	public void addSheetConnect(String targetFile, String targetSheetName, 
			String sourceFile, String sourceSheetName ) throws Exception{
		logger.info("��" + sourceFile + "[" + sourceSheetName + "] д�� " + targetFile + "[" + targetSheetName + "]  ��ʼ ");
		
		Workbook  targetWorkbook = CreateExeclUtil.getReadWorkbook(targetFile);
//		Workbook  sourceWorkbook = CreateExeclUtil.getReadWorkbook(sourceFile);
		Sheet targetSheet = getSheet(targetFile, targetSheetName);
		if(targetSheet == null){
			logger.error("��Sheet����[" + targetSheetName + "]");
			targetSheet = targetWorkbook.createSheet(targetSheetName);
		}
		Sheet sourceSheet = getSheet(sourceFile, sourceSheetName);
		//targetSheet = AddSheeRows(targetSheet, sourceSheet, targetFile);
		targetSheet = AddSheeRows(targetSheet, sourceSheet);
		
		writeWorkbookToFile(targetSheet.getWorkbook(),  targetFile  );
		logger.info("��" + sourceFile + "[" + sourceSheetName + "] д�� " + targetFile + "[" + targetSheetName + "]  ���� ");
		
	}
	
	
	
	public static boolean writeWorkbookToFile(Workbook workbook, String outFile) {
		logger.info("�� workbook д���ļ� " + outFile  + "  ��ʼ ");
		FileOutputStream fOut =  null;
		try {
			File file = new File(outFile);
//			if (file.exists() && file.isFile()) {
//				logger.info(outFile + "�Ѵ���, ��ʼɾ���ļ� ");
//				file.delete();
//				logger.info(outFile + "�Ѵ���, ɾ���ļ���� ");
//			}
			fOut = new FileOutputStream(outFile);
			workbook.write(fOut);
			fOut.flush();  
			fOut.close();  // �����������ر��ļ�  
		} catch (FileNotFoundException e) {
			logger.error("FileNotFoundException", e);
			return false;
		} catch (IOException e) {
			logger.error("IOException", e);
			return false;
		}
		logger.info("�� workbook д���ļ� " + outFile  + "  ���");
		return true;

	}
	
	

}




