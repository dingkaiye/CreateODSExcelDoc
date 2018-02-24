package comm;

import java.util.List;

import org.apache.log4j.LogManager;
import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;

public class ReadExcelUtil {

	private static Logger logger = LogManager.getLogger("Execl"); 
	

	// ��׼���������� 
	public static void CreatOdsType(TableInfo tbInfo) {
		logger.info("����" + tbInfo.getTabEnName() + ":" + tbInfo.getTabChName() + "  ����ת��   ��ʼ" );
		List<DataLineInfo> dataLineList = tbInfo.getDatalineList();
		for (DataLineInfo dataLine : dataLineList) {
			String type = dataLine.getType().toUpperCase().trim();
			String name = dataLine.getChName().toUpperCase().trim();
			String targetType = "";
			if (!"".equals(type) && null != type) {
				if ("DATE".equals(type)) {
					targetType = "VARCHAR(8)";
				} else if ("TIME".equals(type)) {
					targetType = "VARCHAR(6)";
				} else if ("TIMESTAMP".equals(type)) {
					targetType = "VARCHAR(14)";
				} else if ( type.matches("^CHAR") || type.startsWith("CHAR") ) {
					targetType = "VAR" + type;
				} else if ( type.matches("^DECIMAL") || type.startsWith("DECIMAL")) {
					targetType = "NUMBER";
				} else if ( "INTEGER".equals(type) ) {
					targetType = "NUMBER";
				} else if ( type.startsWith("INT") ) {
					targetType = "NUMBER";
				} else {
					targetType = type;
				}

				// �����ת��Ϊ NUMBER(22,2)
				if ( type.contains("DECIMAL") || type.contains("NUMBER")
						|| type.contains("INT") || type.contains("CHAR")
						|| type.contains("VARCHAR") ){
					if (name.endsWith("���")) {
						targetType = "NUMBER(22,2)";
					} else if (name.endsWith("���")) {
						targetType = "NUMBER(22,2)";
					} else if (name.endsWith("������")) {
						targetType = "NUMBER(22,2)";
					} else if (name.endsWith("�ܶ�")) {
						targetType = "NUMBER(22,2)";
					} else if ("�ܶ�".equals(name) || "���".equals(name) || "������".equals(name) || "������".equals(name)) {
						targetType = "NUMBER(22,2)";
					}
				}
			}
			dataLine.setOdsType(targetType);
		}
		logger.info("����" + tbInfo.getTabEnName() + ":" + tbInfo.getTabChName() + " ����ת��   ���" );
	}
	
	public static void CreatTableListOdsType(List<TableInfo> tbInfoList) {
		for(TableInfo tbInfo: tbInfoList){
			CreatOdsType(tbInfo);
		}
	}
	
	
	/**
	 * ��¼ö��ֵ
	 * ���� splits ���ȼ��ָ� 
	 * @param tbInfo
	 * @param splits
	 */
	public static void CreatEnum(TableInfo tbInfo, String[] splits, String[] insidesplits) {
		if(logger.isInfoEnabled()){
			String tabEnName = tbInfo.getTabEnName();
//			String tabChName = tbInfo.getTabChName();
			logger.info("�� " + tbInfo.getTabEnName() + "����¼ö��ֵ��Ϣ  ��ʼ ");
		}
		
		if (splits == null || splits.length == 0) {
			splits = new String[1];
			splits[0] = "\n";
		}

		List<DataLineInfo> dataLineList = tbInfo.getDatalineList();
		for (DataLineInfo dataLine : dataLineList) {
			String desc = dataLine.getDesc();
			String[] enums = null;
			if (desc != null && !"".equals(desc)) {
//				logger.debug("������Ϣ" + desc);
				for (String split : splits) {
					if (desc.contains(split)) {
						enums = desc.split(split);
						for (String curEnum : enums) {
							// �ָ� key �� value
							for (String insplit : insidesplits) { 
								if (curEnum.contains(insplit)) {
									String key = null;
									String value = null;
									String[] tmpStr = curEnum.split(insplit);
									if(tmpStr.length > 1){
										key = curEnum.split(insplit)[0];
										value = curEnum.split(insplit)[1];
										// String key = curEnum;
										// String value = curEnum;
										dataLine.addEnum(key, value);
									}else {
										dataLine.addEnum(curEnum, curEnum);
									}
								}
							}
						}
						break;
					} 
					else {
						continue;
					}
				}
			} 
			else {
				continue;
			}
		}
		
		if(logger.isInfoEnabled()){
			String tabEnName = tbInfo.getTabEnName();
//			String tabChName = tbInfo.getTabChName();
			logger.info("�� " + tbInfo.getTabEnName() + "����¼ö��ֵ��Ϣ  ��� ");
		}
	}
	
	public static void CreatTableListEnum(List<TableInfo> tbInfoList, String[] splits, String[] insplits) {
		for(TableInfo tbInfo: tbInfoList){
			CreatEnum(tbInfo, splits, insplits);
		}
	}
	
	
	/**
	 * ��ȡ��Ԫ������, �����Ԫ��Ϊnull, ���ؿմ�
	 * @param row
	 * @param Col
	 * @return
	 */

	public static String getCellString(Row row, int Col) {
		if(row == null){
			return "";
		}
		String cellStringValue = null;
		CellType cellType = CellType.STRING;
		Cell cell = row.getCell(Col);
		if (row != null && cell != null) {
			cell.setCellType(cellType);
//			logger.debug(row.toString() + "[" + Col + "]:" + cell + " Type:" + cellType);
			cellStringValue = cell.getStringCellValue();
		} else {
//			logger.debug(row.toString() + "[" + Col + "]:" + cell + " NULL ");
			cellStringValue = "";
		}
		return cellStringValue;
	}
}