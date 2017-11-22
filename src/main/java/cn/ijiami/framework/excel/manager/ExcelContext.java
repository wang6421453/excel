package cn.ijiami.framework.excel.manager;

import java.util.List;

/**
 * Excel上下文
 * 
 * @author wl
 *
 */
public class ExcelContext {

	/**
	 * 文件路径
	 */
	private String filePath;

	/**
	 * 字段集合
	 */
	private List<ExcelField> fields;

	/**
	 * 实体类class
	 */
	private Class<?> clazz;

	/**sheet页索引*/
	private int sheetIndex;

	/**sheet页名称*/
	private String sheetName;

	/**标题行索引*/
	private int titleRowIndex;

	/**内容行索引*/
	private int dataRowIndex;

	public String getFilePath() {
		return filePath;
	}

	public void setFilePath(String filePath) {
		this.filePath = filePath;
	}

	public List<ExcelField> getFields() {
		return fields;
	}

	public void setFields(List<ExcelField> fields) {
		this.fields = fields;
	}

	public Class<?> getClazz() {
		return clazz;
	}

	public void setClazz(Class<?> clazz) {
		this.clazz = clazz;
	}

	public int getSheetIndex() {
		return sheetIndex;
	}

	public void setSheetIndex(int sheetIndex) {
		this.sheetIndex = sheetIndex;
	}

	public int getTitleRowIndex() {
		return titleRowIndex;
	}

	public void setTitleRowIndex(int titleRowIndex) {
		this.titleRowIndex = titleRowIndex;
	}

	public int getDataRowIndex() {
		return dataRowIndex;
	}

	public void setDataRowIndex(int dataRowIndex) {
		this.dataRowIndex = dataRowIndex;
	}

	public String getSheetName() {
		return sheetName;
	}

	public void setSheetName(String sheetName) {
		this.sheetName = sheetName;
	}
}
