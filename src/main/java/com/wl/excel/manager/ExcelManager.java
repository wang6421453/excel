package com.wl.excel.manager;

import com.wl.excel.parser.ExcelParser;

import java.util.List;

/**
 * excel管理类
 * 
 * @author wl
 *
 */
public class ExcelManager {

	private static final ExcelManager instance = new ExcelManager();

	public static ExcelManager getInstance() {
		return instance;
	}

	/**
	 * 生成excel上下文
	 * @param filePath excel文件路径
	 * @param clazz 实体类class
	 * @return ExcelContext
	 */
	public ExcelContext generateExcelContext(String filePath, Class<?> clazz){
		ExcelContext excelContext = new ExcelContext();
		excelContext.setFilePath(filePath);
		excelContext.setClazz(clazz);
		ExcelParser.parseClass(clazz, excelContext);
		return excelContext;
	}

	/**
	 * 导入excel为集合
	 * 
	 * @param context
	 * @return List<?>
	 */
	public List<?> importExcel(ExcelContext context) throws Exception{
		return ExcelParser.parseExcel(context);
	}

	/**
	 * 导出数据到指定文件
	 * 
	 * @param excelContext
	 */
	public void exportExcel(ExcelContext excelContext, List<?> listData) throws Exception{
		ExcelParser.exportExcel(excelContext, listData);
	}

}
