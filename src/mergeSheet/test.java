package mergeSheet;

public class test {

	
	public static void main(String agrv[]) throws Exception{
		
		// д���ļ� 
		String tergetfile = "D:\\Temp\\AML\\All.xls";
		
		MergeSheet ms = new MergeSheet();
		
		
	//	ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls", "Դ�ֶ�", "D:\\Temp\\AML\\ODS��׼���ݽṹ�ĵ�(����㼰ϵͳ��¼��)-AGL.xls", "Դ�ֶ�");
		//ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls", "Դ�ֶ�", "D:\\Temp\\AML\\ODS��׼���ݽṹ�ĵ�(����㼰ϵͳ��¼��)-ATM.xls", "Դ�ֶ�");
		ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls","Դ�ֶ�","D:\\Temp\\AML\\ODS��׼���ݽṹ�ĵ�(����㼰ϵͳ��¼��)-ECIF.xls","Դ�ֶ�");
		ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls","Դ�ֶ�","D:\\Temp\\AML\\ODS��׼���ݽṹ�ĵ�(����㼰ϵͳ��¼��)-POS.xls","Դ�ֶ�");
		ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls","Դ�ֶ�","D:\\Temp\\AML\\ODS��׼���ݽṹ�ĵ�(����㼰ϵͳ��¼��)-TFS.xls","Դ�ֶ�");
		ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls","Դ�ֶ�","D:\\Temp\\AML\\ODS��׼���ݽṹ�ĵ�(����㼰ϵͳ��¼��)-HR.xls","Դ�ֶ�");
		ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls","Դ�ֶ�","D:\\Temp\\AML\\ODS��׼���ݽṹ�ĵ�(����㼰ϵͳ��¼��)-AGL.xls","Դ�ֶ�");
		ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls","Դ�ֶ�","D:\\Temp\\AML\\ODS��׼���ݽṹ�ĵ�(����㼰ϵͳ��¼��)-UPB.xls","Դ�ֶ�");
		ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls","Դ�ֶ�","D:\\Temp\\AML\\ODS��׼���ݽṹ�ĵ�(����㼰ϵͳ��¼��)-OBM.xls","Դ�ֶ�");
		ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls","Դ�ֶ�","D:\\Temp\\AML\\ODS��׼���ݽṹ�ĵ�(����㼰ϵͳ��¼��)-TBC.xls","Դ�ֶ�");
		ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls","Դ�ֶ�","D:\\Temp\\AML\\ODS��׼���ݽṹ�ĵ�(����㼰ϵͳ��¼��)-ECD.xls","Դ�ֶ�");
		ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls","Դ�ֶ�","D:\\Temp\\AML\\ODS��׼���ݽṹ�ĵ�(����㼰ϵͳ��¼��)-CMS.xls","Դ�ֶ�");
		ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls","Դ�ֶ�","D:\\Temp\\AML\\ODS��׼���ݽṹ�ĵ�(����㼰ϵͳ��¼��)-ATM.xls","Դ�ֶ�");
		
		ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls","Դ�ֶ�","D:\\Temp\\AML\\ODS��׼���ݽṹ�ĵ�(����㼰ϵͳ��¼��)-CBS-�����ʲ��ֲ�.xls","Դ�ֶ�");
		ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls","Դ�ֶ�","D:\\Temp\\AML\\ODS��׼���ݽṹ�ĵ�(����㼰ϵͳ��¼��)-CBS-���˽��ۻ�&��������ֲ�.xls","Դ�ֶ�");
		ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls","Դ�ֶ�","D:\\Temp\\AML\\ODS��׼���ݽṹ�ĵ�(����㼰ϵͳ��¼��)-CBS-���ʷֲ�.xls","Դ�ֶ�");
		ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls","Դ�ֶ�","D:\\Temp\\AML\\ODS��׼���ݽṹ�ĵ�(����㼰ϵͳ��¼��)-CBS-��Χ�ֲ�.xls","Դ�ֶ�");
		ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls","Դ�ֶ�","D:\\Temp\\AML\\ODS��׼���ݽṹ�ĵ�(����㼰ϵͳ��¼��)-CBS-�����ֲ�.xls","Դ�ֶ�");
		ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls","Դ�ֶ�","D:\\Temp\\AML\\ODS��׼���ݽṹ�ĵ�(����㼰ϵͳ��¼��)-CBS-��Ա�ֲ�.xls","Դ�ֶ�");
		ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls","Դ�ֶ�","D:\\Temp\\AML\\ODS��׼���ݽṹ�ĵ�(����㼰ϵͳ��¼��)-CBS-�ֽ�ֲ�.xls","Դ�ֶ�");
		ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls","Դ�ֶ�","D:\\Temp\\AML\\ODS��׼���ݽṹ�ĵ�(����㼰ϵͳ��¼��)-CBS-����ֲ�.xls","Դ�ֶ�");
		ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls","Դ�ֶ�","D:\\Temp\\AML\\ODS��׼���ݽṹ�ĵ�(����㼰ϵͳ��¼��)-CBS-�ɽ�ֲ�.xls","Դ�ֶ�");
		ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls","Դ�ֶ�","D:\\Temp\\AML\\ODS��׼���ݽṹ�ĵ�(����㼰ϵͳ��¼��)-CBS-���÷ֲ�.xls","Դ�ֶ�");
		ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls","Դ�ֶ�","D:\\Temp\\AML\\ODS��׼���ݽṹ�ĵ�(����㼰ϵͳ��¼��)-CBS-�ʲ�֤ȯ���ֲ�.xls","Դ�ֶ�");
		ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls","Դ�ֶ�","D:\\Temp\\AML\\ODS��׼���ݽṹ�ĵ�(����㼰ϵͳ��¼��)-CBS-�ʽ�ֲ�.xls","Դ�ֶ�");
		ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls","Դ�ֶ�","D:\\Temp\\AML\\ODS��׼���ݽṹ�ĵ�(����㼰ϵͳ��¼��)-CBS_��Ʒֲ�.xls","Դ�ֶ�");
		ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls","Դ�ֶ�","D:\\Temp\\AML\\ODS��׼���ݽṹ�ĵ�(����㼰ϵͳ��¼��)-CBS_ƾ֤�ֲ�.xls","Դ�ֶ�");
		ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls","Դ�ֶ�","D:\\Temp\\AML\\ODS��׼���ݽṹ�ĵ�(����㼰ϵͳ��¼��)-CBS_���ֲ�.xls","Դ�ֶ�");
		ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls","Դ�ֶ�","D:\\Temp\\AML\\ODS��׼���ݽṹ�ĵ�(����㼰ϵͳ��¼��)-CBS_����ֲ�.xls","Դ�ֶ�");
		ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls","Դ�ֶ�","D:\\Temp\\AML\\ODS��׼���ݽṹ�ĵ�(����㼰ϵͳ��¼��)-CBS_ͨ�÷ֲ�.xls","Դ�ֶ�");
		ms.addSheetConnect("D:\\Temp\\AML\\ALL.xls","Դ�ֶ�","D:\\Temp\\AML\\ODS��׼���ݽṹ�ĵ�(����㼰ϵͳ��¼��)-CBS_���п��ֲ�.xls","Դ�ֶ�");
	
	}
	
	
	
	
	
	
}