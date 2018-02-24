package comm;

import java.util.ArrayList;
import java.util.List;

public class DataLineInfo {
	private int     seqNo           ;     // 序号
	private String  enName          ;     // 字段英文名	
	private String  chName          ;     // 字段中文名	
	private String  type            ;     // 字段类型（含长度）	
	private String  isNull          ;     // 是否允许为空	
	private String  isPK            ;     // 主键	
	private String  isFK            ;     // 外键/索引	
	private String  defaultValue    ;     // 缺省值	
	private String  desc            ;     // 字段说明	
	private String  dataStandardNo  ;     // 数据标准编号	
	private String  dataStandardName;     // 数据标准中文名称
	private String  odsType ;             // 目标数据类型 
	private List<ColumnEnum> enums = null ;      // 枚举值 
	
	
	
	
	public List<ColumnEnum> getEnums() {
		return enums;
	}
	public void setEnums(List<ColumnEnum> enums) {
		this.enums = enums;
	}
	public int getSeqNo() {
		return seqNo;
	}
	public void setSeqNo(int seqNo) {
		this.seqNo = seqNo;
	}
	public String getEnName() {
		return enName;
	}
	public String getChName() {
		return chName;
	}
	public String getType() {
		return type;
	}
	public String getIsNull() {
		return isNull;
	}
	public String getIsPK() {
		return isPK;
	}
	public String getIsFK() {
		return isFK;
	}
	public String getDefaultValue() {
		return defaultValue;
	}
	public String getDesc() {
		return desc;
	}
	public String getDataStandardNo() {
		return dataStandardNo;
	}
	public String getDataStandardName() {
		return dataStandardName;
	}
	public void setEnName(String enName) {
		this.enName = enName;
	}
	public void setChName(String chName) {
		this.chName = chName;
	}
	public void setType(String type) {
		this.type = type;
	}
	public void setIsNull(String isNull) {
		this.isNull = isNull;
	}
	public void setIsPK(String isPK) {
		this.isPK = isPK;
	}
	public void setIsFK(String isFK) {
		this.isFK = isFK;
	}
	public void setDefaultValue(String defaultValue) {
		this.defaultValue = defaultValue;
	}
	public void setDesc(String desc) {
		this.desc = desc;
	}
	public void setDataStandardNo(String dataStandardNo) {
		this.dataStandardNo = dataStandardNo;
	}
	public void setDataStandardName(String dataStandardName) {
		this.dataStandardName = dataStandardName;
	}
	public String getOdsType() {
		return odsType;
	}
	public void setOdsType(String odsType) {
		this.odsType = odsType;
	}
	
	public void addEnum(String key, String value){
		if(enums == null){
			this.enums = new ArrayList<ColumnEnum>();
		}
		this.enums.add(new ColumnEnum(key, value)) ;
	}
	
	public void removeEnum(String key, String value){
		if(enums == null){
			this.enums = new ArrayList<ColumnEnum>();
		}
		int index = 0;
		for(ColumnEnum ce : enums){
			if(ce.getKey().equals(key)){
				this.enums.remove(index);
			}
			index ++;
		}
	}
	
	
	
}
