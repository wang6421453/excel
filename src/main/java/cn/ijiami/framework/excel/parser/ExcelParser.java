package cn.ijiami.framework.excel.parser;

import cn.ijiami.framework.excel.constants.TypeConstants;
import cn.ijiami.framework.excel.manager.ExcelContext;
import cn.ijiami.framework.excel.manager.ExcelField;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

/**
 * excel解析器
 *
 * @author wl
 * @date 2017/11/8
 */
public class ExcelParser {

	/**
	 * 解析字段
	 * @param clazz 实体类class
	 * @return List<ExcelField>(字段集合)
	 */
	public static void parseClass(Class<?> clazz, ExcelContext excelContext){
		//解析类说明注解
		Excel excel = clazz.getAnnotation(Excel.class);
		if(excel != null){
			int sheetIndex = excel.sheetIndex() - 1;
			excelContext.setSheetIndex(sheetIndex);
			int titleRowIndex = excel.titleRowIndex() - 1;
			excelContext.setTitleRowIndex(titleRowIndex);
			int dataRowIndex = excel.dataRowIndex() - 1;
			excelContext.setDataRowIndex(dataRowIndex);
		}


		//解析field
		List<ExcelField> lstField = new ArrayList<>();
		Field[] fields = clazz.getDeclaredFields();
		for(Field filed : fields){
			ExcelColumn excelColumn = filed.getAnnotation(ExcelColumn.class);
			if(excelColumn != null){
				int columnIndex = excelColumn.value() - 1;
				String fieldName = filed.getName();
				ExcelField excelField = new ExcelField();
				excelField.setColumnIndex(columnIndex);
				excelField.setFieldName(fieldName);
				lstField.add(excelField);
			}
		}
		excelContext.setFields(lstField);
	}

	/**
	 * 解析excel
	 * @param excelContext excelContext
	 * @return List
	 */
	public static List<?> parseExcel(ExcelContext excelContext) throws Exception{
		List list = new ArrayList();
		//文件路径
		String filePath = excelContext.getFilePath();
		//获取字段集合
		List<ExcelField> lstField = excelContext.getFields();
		//实体类class
		Class clazz = excelContext.getClazz();
		File xlsFile = new File(filePath);
		//获得工作薄对象
		Workbook workbook = WorkbookFactory.create(xlsFile);
		//获得工作表索引
		int sheetIndex = excelContext.getSheetIndex();
		Sheet sheet = workbook.getSheetAt(sheetIndex);
		//获得行数
		int rowCount = sheet.getLastRowNum() + 1;
		//获取数据起始行
		int dataRowIndex = excelContext.getDataRowIndex();
		//读取数据
		for(int rowIndex = dataRowIndex; rowIndex < rowCount; rowIndex++){
			Row row = sheet.getRow(rowIndex);
			Object obj = clazz.newInstance();
			for(ExcelField excelField : lstField){
				int columnIndex = excelField.getColumnIndex();
				Cell cell = row.getCell(columnIndex);
				String fieldName = excelField.getFieldName();
				Field field = clazz.getDeclaredField(fieldName);
				setFieldValue(field, cell, obj);
			}
			list.add(obj);
		}
		return list;
	}

	/**
	 * 根据字段类型设置字段值
	 * @param field 字段
	 * @param cell excel单元格
	 * @param obj 传入类的实例
	 * @throws Exception
	 */
	private static void setFieldValue(Field field, Cell cell, Object obj) throws Exception{
		String fieldType = field.getType().toString();
		fieldType = fieldType.substring(fieldType.lastIndexOf(".") + 1).toUpperCase();
		field.setAccessible(true);
		if(TypeConstants.STRING.toUpperCase().startsWith(fieldType)){
			field.set(obj, cell.getStringCellValue());
		}else if(TypeConstants.DOUBLE.toUpperCase().startsWith(fieldType)){
			field.set(obj, cell.getNumericCellValue());
		}else if(TypeConstants.INTEGER.toUpperCase().startsWith(fieldType)){
			field.set(obj, new Double(cell.getNumericCellValue()).intValue());
		}else if(TypeConstants.DATE.toUpperCase().startsWith(fieldType)){
			field.set(obj, cell.getDateCellValue());
		}else if(TypeConstants.BOOLEAN.toUpperCase().startsWith(fieldType)){
			field.set(obj, cell.getBooleanCellValue());
		}else if(TypeConstants.FLOAT.toUpperCase().startsWith(fieldType)){
			field.set(obj, new Double(cell.getNumericCellValue()).floatValue());
		}else if(TypeConstants.LONG.toUpperCase().startsWith(fieldType)){
			field.set(obj, new Double(cell.getNumericCellValue()).longValue());
		}else if(TypeConstants.SHORT.toUpperCase().startsWith(fieldType)){
			field.set(obj, new Double(cell.getNumericCellValue()).shortValue());
		}else if(TypeConstants.BYTE.toUpperCase().startsWith(fieldType)){
			field.set(obj, new Double(cell.getNumericCellValue()).byteValue());
		}else if(TypeConstants.CHAR.toUpperCase().startsWith(fieldType)){
			String cellValue = cell.getStringCellValue();
			if(StringUtils.isNotEmpty(cellValue)){
				field.set(obj, cell.getStringCellValue().charAt(0));
			}else{
				field.set(obj, "");
			}
		}

	}

	/**
	 * 导出excel
	 * @param excelContext
	 * @param listData
	 */
	public static void exportExcel(ExcelContext excelContext, List<?> listData) throws Exception{
		//文件路径
		String filePath = excelContext.getFilePath();
		//获取字段集合
		List<ExcelField> lstField = excelContext.getFields();
		//实体类class
		Class clazz = excelContext.getClazz();
		//获得工作表索引
		int sheetIndex = excelContext.getSheetIndex();
		//获取数据起始行
		int dataRowIndex = excelContext.getDataRowIndex();
		//数据总数
		int dataCount = listData.size();
		//创建工作薄
		Workbook workbook = WorkbookFactory.create(new FileInputStream(filePath));
		//获取工作表
		Sheet sheet = workbook.getSheetAt(sheetIndex);
		for(int rowIndex = dataRowIndex; rowIndex < dataCount + dataRowIndex ; rowIndex++){
			Row row = sheet.createRow(rowIndex);
			Object data = listData.get(rowIndex - dataRowIndex);
			for(ExcelField excelField : lstField){
				int columnIndex = excelField.getColumnIndex();
				String fieldName = excelField.getFieldName();
				Field field = clazz.getDeclaredField(fieldName);
				field.setAccessible(true);
				row.createCell(columnIndex).setCellValue(field.get(data).toString());
			}
		}

		workbook.write(new FileOutputStream(filePath));
	}
}
