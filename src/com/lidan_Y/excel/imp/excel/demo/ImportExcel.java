package com.lidan_Y.excel.imp.excel.demo;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * 
 * 获取导入的excel文件的数据
 * 
 * @author ilidan_Y
 *
 */
public class ImportExcel {

	private final Integer HEAD = 0;
	private final Integer START = 1;
	private Map<Integer, String> headMap;
	private Map<String, String> relationMap;
	private Logger log = Logger.getLogger(this.getClass());
	private InputStream inputStream = null;

	public ImportExcel() {
	}

	/**
	 * excel表格head和po之间映射信息
	 * 
	 * @param map
	 *            head对应实体关系表(key:excel head单元格值，value：对应excel head的po属性)
	 * @param inputStream
	 *            excel文件流
	 */
	public ImportExcel(Map<String, String> map, InputStream inputStream) {
		this.relationMap = map;
		this.inputStream = inputStream;
	}

	/**
	 * 获取excel文件的数据 excel2007版本
	 * 
	 * @return
	 * @throws IOException
	 */
	public <T> List<T> getExcelDataList(Class<T> clazz) throws Exception {
		List<T> list = new ArrayList<T>();
		// 创建一个空的工作薄
		XSSFWorkbook wb = null;
		try {
			wb = new XSSFWorkbook(inputStream);
		} catch (IOException e) {
			e.getMessage();
			throw new IOException("Can not create XSSFWorkbook!");
		}
		// 创建一个sheet
		XSSFSheet sheet = wb.getSheetAt(0);
		initExcelHeadData(sheet);
		int rowNum = sheet.getLastRowNum();
		XSSFRow row = null;
		for (int i = START; i <= rowNum; i++) {// 从excel第二行开始度数据
			row = sheet.getRow(i);
			T instance = null;
			try {
				instance = clazz.newInstance();
			} catch (InstantiationException e) {
				e.printStackTrace();
			} catch (IllegalAccessException e) {
				e.printStackTrace();
			}
			if (null != row) {
				XSSFCell cell = null;
				int cellNum = row.getLastCellNum();
				for (int j = 0; j < cellNum; j++) {
					cell = row.getCell(j);
					String value = getCellValue(cell);
					if (null != value && !"".equals(value)) {
						setValueToInstance(clazz, instance, relationMap.get(headMap.get(j)), value);
					} else {

					}
				}

			} else {
				int r = i + 1;
				log.debug("第" + r + "行是空行！");
			}
			list.add(instance);
		}

		return list;
	}

	/**
	 * 注入属性值
	 * 
	 * @param clazz
	 * @param instance
	 * @param methodName
	 * @param value
	 * @throws Exception
	 */
	private <T> void setValueToInstance(Class<T> clazz, T instance, String methodName, String value) throws Exception {
		String getMethod = ReflectionUtils.createGetMethodName(methodName);
		String setMethod = ReflectionUtils.createSetMethodName(methodName);
		Class<T> type = (Class<T>) clazz.getDeclaredMethod(getMethod, null).getReturnType();
		Method method = clazz.getMethod(setMethod, type);
		if (type == String.class) {
			method.invoke(instance, value);
		} else if (type == int.class || type == Integer.class) {
			method.invoke(instance, Integer.parseInt(value));

		} else if (type == long.class || type == Long.class) {
			method.invoke(instance, Long.parseLong(value));

		} else if (type == float.class || type == Float.class) {
			method.invoke(instance, Float.parseFloat(value));

		} else if (type == double.class || type == Double.class) {
			method.invoke(instance, Double.parseDouble(value));

		} else if (type == Date.class) {
			method.invoke(instance, new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(value));
		}
	}

	/**
	 * 将单元格的数据转化为String
	 * 
	 * @param cell
	 * @return
	 */
	private String getCellValue(XSSFCell cell) {
		String value = null;
		SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
		switch (cell.getCellType()) {
		case XSSFCell.CELL_TYPE_BOOLEAN:
			value = String.valueOf(cell.getBooleanCellValue());
			break;
		case XSSFCell.CELL_TYPE_NUMERIC:
			// 判断当前的cell是否为Date
			if (HSSFDateUtil.isCellDateFormatted(cell)) {
				value = df.format(cell.getDateCellValue());
			} else {
				value = String.valueOf((long) cell.getNumericCellValue());
			}
			break;
		case XSSFCell.CELL_TYPE_STRING:
			value = cell.getStringCellValue();
			break;
		case XSSFCell.CELL_TYPE_FORMULA:
			log.debug("不支持函数！");
			break;
		}

		return value;
	}

	/**
	 * 初始化excel head 标题
	 * 
	 * @param sheet
	 */
	private void initExcelHeadData(XSSFSheet sheet) {
		this.headMap = new HashMap<Integer, String>();
		XSSFRow head = sheet.getRow(HEAD);
		int cellNum = head.getLastCellNum();
		XSSFCell cell = null;
		for (int i = 0; i < cellNum; i++) {
			cell = head.getCell(i);
			headMap.put(i, cell.getStringCellValue());
		}
	}

//	public static void main(String[] args) {
//		String filePath = "/Users/gemii.yangyang/Desktop/test.xlsx";
//		File file = new File(filePath);
//		FileInputStream inputStream = null;
//		try {
//			inputStream = new FileInputStream(file);
//		} catch (FileNotFoundException e) {
//			e.printStackTrace();
//		}
//		Map<String, String> map = new HashMap<String, String>();
//		map.put("roomID", "roomId");
//		map.put("roomName", "roomName");
//		map.put("username", "username");
//		map.put("city", "city");
//		map.put("memberNums", "memberNums");
//		ImportExcel importExcel = new ImportExcel(map, inputStream);
//		try {
//			List<ExcelTemplate> list = importExcel.getExcelDataList(ExcelTemplate.class);
//			System.out.println(list);
//		} catch (Exception e) {
//			// TODO Auto-generated catch block
//			e.printStackTrace();
//		}
//	}

}
