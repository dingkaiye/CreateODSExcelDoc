package cbs;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.log4j.LogManager;
import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import comm.DataLineInfo;
import comm.ReadExcelUtil;
import comm.CreateExeclUtil;
import comm.TableInfo;

public class CreateCBSExecl {
	private static Logger logger = LogManager.getLogger("Execl");   
	public static void main(String[] argv){
		if(argv.length != 4){
			logger.error("������������\n");
			logger.error("Դϵͳ�����ֵ��ļ�   ģ���ļ�    ����ļ�  ģ������");
			System.exit(-1);
		}
		String fileName = argv[0];  
		String mufileName = argv[1];  
		String outFile = argv[2];   
		String muName = argv[3];
	
		boolean errorFlg = false;
		if(fileName == null || "".equals(fileName)){
			logger.error("�����ֵ��ļ�����Ϊ��\n");
			errorFlg = true;
		}
		if(mufileName == null || "".equals(mufileName)){
			logger.error("ģ���ļ�����Ϊ��\n");
			errorFlg = true;
		}
		if(outFile == null || "".equals(outFile)){
			logger.error("����ļ� ����Ϊ��\n");
			errorFlg = true;
		}
		
		File file=new File(fileName);
        if(!file.exists() || !file.isFile()){
        	logger.error("�����ֵ��ļ�: " + fileName + " ������");
        	errorFlg = true;
		}
        file=new File(mufileName);
        if(!file.exists() || !file.isFile()){
        	logger.error("ģ���ļ�        : " + mufileName + " ������");
        	errorFlg = true;
		}

		if(errorFlg){
			logger.error("   ����!!!");
			System.exit(-1);
		}
		
		logger.debug("�����ֵ��ļ�:" + fileName);
		logger.debug("ģ���ļ�        :" + mufileName);
		logger.debug("����ļ�         :" + outFile);
		
		ReadExeclSheet worker = new ReadExeclSheet();
		
		Workbook  workbook = CreateExeclUtil.getReadWorkbook(fileName);
		List<TableInfo> tablelist = worker.readTableList(workbook, "����ͼ)�嵥");
		ReadExeclSheet readExeclSheet = new ReadExeclSheet ();
		for(TableInfo tableInfoCur : tablelist) {
			String tabEnName = tableInfoCur.getTabEnName();
			String tabChName = tableInfoCur.getTabChName();
			
			String SheetName = tableInfoCur.getLink();
			
			List<DataLineInfo> dataList = readExeclSheet.readTableDef ( workbook, SheetName, tabEnName,  tabChName ) ; 
			tableInfoCur.setDatalineList(dataList);
		}
		ReadExcelUtil.CreatTableListOdsType(tablelist);
		String[] splits = {"\n", ";", "-"};
		ReadExcelUtil.CreatTableListEnum(tablelist, splits, splits);
		
		Workbook  resultWorkbook = CreateExeclUtil.getReadWorkbook(mufileName);
		Sheet emenuSheet = resultWorkbook.getSheet("Դϵͳ����");
		Sheet tableListSheet = resultWorkbook.getSheet("Դ��");
		Sheet ableDefSheet = resultWorkbook.getSheet("Դ�ֶ�");
		
		CreateExeclUtil.CreatTableDefSheet("CBS", ableDefSheet, tablelist);
		CreateExeclUtil.CreatTableListSheet("CBS", muName, tableListSheet, tablelist );
		CreateExeclUtil.CreatEmenuSheet("CBS", emenuSheet, tablelist);
		
		CreateExeclUtil.writeWorkbookToFile(resultWorkbook, outFile);
		
		logger.debug(tablelist);
	}

}
