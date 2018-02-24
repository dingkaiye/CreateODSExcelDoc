package comm;

import java.util.ArrayList;
import java.util.List;

public class TableInfo {
	
	private int  seqNo      ;    // 序号	
	private String  Model      ;    // 模块	
	private String  tabEnName  ;    // 表英文名称	
	private String  tabChName  ;    // 表中文名称	
	private String  tabDesc    ;    // 表功能描述	
	private String  tabOrView  ;    // 表/视图	
	private String  version    ;    // 版本	
	private String  altTime    ;    // 变更时间
	private List<DataLineInfo> DatalineList ;
	private String link ; // 链接
	
	
	public String getLink() {
		return link;
	}


	public void setLink(String link) {
		this.link = link;
	}


	public TableInfo () {
		seqNo      = -1 ;
		Model      = new String () ;
		tabEnName  = new String () ;
		tabChName  = new String () ;
		tabDesc    = new String () ;
		tabOrView  = new String () ;
		version    = new String () ;
		altTime    = new String () ;
		DatalineList = new ArrayList<DataLineInfo>();
	}
	
	
	public int getSeqNo() {
		return seqNo;
	}
	public String getModel() {
		return Model;
	}
	public String getTabEnName() {
		return tabEnName;
	}
	public String getTabChName() {
		return tabChName;
	}
	public String getTabDesc() {
		return tabDesc;
	}
	public String getTabOrView() {
		return tabOrView;
	}
	public String getVersion() {
		return version;
	}
	public String getAltTime() {
		return altTime;
	}
	public void setSeqNo(int seqNo) {
		this.seqNo = seqNo;
	}
	public void setModel(String model) {
		Model = model;
	}
	public void setTabEnName(String tabEnName) {
		this.tabEnName = tabEnName;
	}
	public void setTabChName(String tabChName) {
		this.tabChName = tabChName;
	}
	public void setTabDesc(String tabDesc) {
		this.tabDesc = tabDesc;
	}
	public void setTabOrView(String tabOrView) {
		this.tabOrView = tabOrView;
	}
	public void setVersion(String version) {
		this.version = version;
	}
	public void setAltTime(String altTime) {
		this.altTime = altTime;
	}
	public List<DataLineInfo> getDatalineList() {
		return DatalineList;
	}
	public void setDatalineList(List<DataLineInfo> DatalineList) {
		this.DatalineList = DatalineList;
	}

}
