package com.james.poi_util;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.beanutils.BeanUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

@SuppressWarnings("rawtypes")
public class ExcelUtil {
	private static ExcelUtil excelUtil = null;

	private ExcelUtil() {}

	public static ExcelUtil getInstance() {
		if (excelUtil == null) {
			excelUtil = new ExcelUtil();
		}
		return excelUtil;
	}

	/**
	 * 从文件系统中读取Excel文件，封装到对象
	 * 
	 * @param path
	 *            文件系统中的绝对路径
	 * @param clazz
	 *            将Excel内容封装的对象
	 * @param readLine
	 *            从第几行开始读取数据
	 * @param tailLine
	 *            尾部有多少行非clazz对象数据
	 */
	public List<Object> readExcel2ObjectByPath(String path, Class clazz, int readLine, int tailLine) {
		Workbook workbook = null;
		try {
			workbook = WorkbookFactory.create(new File(path));
			return handlerExcel2Object(workbook, clazz, readLine, tailLine);
		} catch (InvalidFormatException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

		return null;
	}

	/**
	 * 从ClassPath中读取Excel文件，封装到对象
	 * 
	 * @param path
	 *            ClassPath
	 * @param clazz
	 *            要封装的对象
	 * @param readLine
	 *            从第几行开始读取数据
	 * @param tailLine
	 *            尾部有多少行非clazz对象数据
	 */
	public List<Object> readExcel2ObjectByClassPath(String path, Class clazz, int readLine, int tailLine) {
		Workbook workbook = null;
		try {
			workbook = WorkbookFactory.create(ExcelUtil.class.getResourceAsStream(path));
			return handlerExcel2Object(workbook, clazz, readLine, tailLine);
		} catch (InvalidFormatException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		return null;
	}

	/**
	 * 从第0行开始读取文件系统中的Excel文件，封装到对象
	 * 
	 * @param path
	 *            文件系统中的绝对路径
	 * @param clazz
	 *            将Excel内容封装的对象
	 */
	public List<Object> readExcel2ObjectByPath(String path, Class clazz) {
		return this.readExcel2ObjectByPath(path, clazz, 0, 0);
	}

	/**
	 * 从第0行开始读取ClassPath中的Excel文件，封装到对象
	 * 
	 * @param path
	 *            ClassPath
	 * @param clazz
	 *            要封装的对象
	 */
	public List<Object> readExcel2ObjectByClassPath(String path, Class clazz) {
		return this.readExcel2ObjectByPath(path, clazz, 0, 0);
	}

	/**
	 * 无模板文件，将对象输出到Excel文件
	 * 
	 * @param outPath
	 *            输出路径
	 * @param objects
	 *            要输出的对象List列表
	 * @param clazz
	 *            要输出的类类型
	 * @param isXssf
	 *            是否XSSFWorkbook
	 */
	public void exportObject2Excel(String outPath, List<Object> objects, Class clazz, boolean isXssf) {
		Workbook workbook = handlerObject2Excel(objects, clazz, isXssf);
		FileOutputStream fileOutputStream = null;
		try {
			fileOutputStream = new FileOutputStream(outPath);
			workbook.write(fileOutputStream);

		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		finally {
			if (fileOutputStream != null) {
				try {
					fileOutputStream.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
	}

	/**
	 * 无模板文件，将对象输出到流
	 * 
	 * @param outputStream
	 *            需要输出的流
	 * @param objects
	 *            要输出的对象List列表
	 * @param clazz
	 *            要输出的类类型
	 * @param isXssf
	 *            是否XSSFWorkbook
	 */
	public void exportObject2Excel(OutputStream outputStream, List<Object> objects, Class clazz, boolean isXssf) {
		Workbook workbook = handlerObject2Excel(objects, clazz, isXssf);
		try {
			workbook.write(outputStream);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	/**
	 * 使用模板文件的样式将对象输出到OutputStream流
	 * 
	 * @param datas
	 *            固定字段的Map
	 * @param inPath
	 *            模板文件的路径
	 * @param outputStream
	 *            输出流
	 * @param objects
	 *            要输出的对象List列表
	 * @param clazz
	 *            要输出的哪一个类对象
	 * @param isClasspath
	 *            是否ClassPath
	 */
	public void exportObject2ExcelByTemplate(Map<String, String> datas, String inPath, OutputStream outputStream, List<Object> objects, Class clazz,
			boolean isClasspath) {
		ExcelTemplate excelTemplate = handlerObject2Excel(datas, inPath, objects, clazz, isClasspath);
		// 写入到outputStream指定的流
		excelTemplate.writeToStream(outputStream);
	}

	/**
	 * 使用模板文件的样式将对象输出到Excel文件
	 * 
	 * @param datas
	 *            固定字段的Map
	 * @param inPath
	 *            模板文件的路径
	 * @param outPath
	 *            输出文件的绝对路径
	 * @param objects
	 *            要输出的对象List列表
	 * @param clazz
	 *            要输出的哪一个类对象
	 * @param isClasspath
	 *            是否ClassPath
	 */
	public void exportObject2ExcelByTemplate(Map<String, String> datas, String inPath, String outPath, List<Object> objects, Class clazz, boolean isClasspath) {
		ExcelTemplate excelTemplate = handlerObject2Excel(datas, inPath, objects, clazz, isClasspath);
		// 写入到outPath指定的文件中
		excelTemplate.writeToFile(outPath);
	}

	/**
	 * 处理Excel文件到JavaBean
	 * 
	 * @param workbook
	 *            需处理的Workbook
	 * @param clazz
	 *            封装的对象
	 * @param readLine
	 *            从第几行开始读
	 * @param tailLine
	 *            尾部有几行不需要读
	 * @return List<Object> 封装对象的List集合
	 */
	private List<Object> handlerExcel2Object(Workbook workbook, Class clazz, int readLine, int tailLine) {
		Sheet sheet = workbook.getSheetAt(0);
		Row row = sheet.getRow(readLine);
		List<Object> objects = new ArrayList<Object>();
		Map<Integer, String> maps = getHeaderMap(row, clazz);
		if (maps == null || maps.size() <= 0) {
			throw new RuntimeException("要读取的Excel文件格式不正确，请检查！");
		}
		Object object = null;
		for (int i = readLine + 1; i <= sheet.getLastRowNum() - tailLine; i++) {
			try {
				row = sheet.getRow(i);
				object = clazz.newInstance();
				for (Cell cell : row) {
					int cellIndex = cell.getColumnIndex();
					String methodName = maps.get(cellIndex).substring(3);
					methodName = methodName.substring(0, 1).toLowerCase() + methodName.substring(1);
					BeanUtils.copyProperty(object, methodName, this.getCellValue(cell));
				}
				objects.add(object);
			} catch (InstantiationException e) {
				e.printStackTrace();
			} catch (IllegalAccessException e) {
				e.printStackTrace();
			} catch (InvocationTargetException e) {
				e.printStackTrace();
			}
		}
		return objects;
	}

	/**
	 * 将Cell的值全部转换成String类型
	 * 
	 * @param cell
	 *            需取值的Cell
	 * @return String 返回转换后的String类型的Cell值
	 */
	private String getCellValue(Cell cell) {
		String object = null;
		switch (cell.getCellType()) {
			case Cell.CELL_TYPE_BLANK:
				object = "";
				break;
			case Cell.CELL_TYPE_BOOLEAN:
				object = String.valueOf(cell.getBooleanCellValue());
				break;
			case Cell.CELL_TYPE_FORMULA:
				object = cell.getCellFormula();
				break;
			case Cell.CELL_TYPE_NUMERIC:
				object = cell.getNumericCellValue() + "";
				break;
			case Cell.CELL_TYPE_STRING:
				object = cell.getStringCellValue();
				break;
			default:
				object = null;
				break;
		}
		return object;
	}

	/**
	 * 根据模板文件输出对象到文件或流
	 * 
	 * @param datas
	 * @param inPath
	 * @param objects
	 * @param clazz
	 * @param isClasspath
	 * @return
	 */
	private ExcelTemplate handlerObject2Excel(Map<String, String> datas, String inPath, List<Object> objects, Class clazz, boolean isClasspath) {
		ExcelTemplate excelTemplate = ExcelTemplate.getInstance();
		try {
			if (isClasspath) {
				// 如果是ClassPath
				excelTemplate.readTemplateByClasspath(inPath);
			} else {
				// 如果是绝对路径
				excelTemplate.readTemplateByPath(inPath);
			}
			List<ExcelHeader> headers = getHeaderList(clazz);
			Collections.sort(headers);
			// 输出标题
			excelTemplate.createNewRow();
			for (ExcelHeader excelHeader : headers) {
				// 创建标题列名
				excelTemplate.createCell(excelHeader.getTitle());
			}

			// 基于BeanUtils输出值
			for (Object object : objects) {
				excelTemplate.createNewRow();
				for (ExcelHeader excelHeader : headers) {
					// 创建Excel文档的内容
					excelTemplate.createCell(BeanUtils.getProperty(object, getMethodName(excelHeader)));
				}
			}

			/*
			 * //基于反射的输出值 for (Object object : objects) { excelTemplate.createNewRow(); for (ExcelHeader excelHeader : headers) { //通过ExcelHeader取得方法名 String
			 * methodName = excelHeader.getMethodName(); //取得方法 Method method = clazz.getDeclaredMethod(methodName); //执行方法 Object rel = method.invoke(object);
			 * //创建Excel文档的内容 excelTemplate.createCell(rel); } }
			 */

			// 替换固定的Title数据
			excelTemplate.replaceConstantData(datas);
		} catch (SecurityException e) {
			throw new RuntimeException("安全性异常！" + e.getMessage());
		} catch (NoSuchMethodException e) {
			throw new RuntimeException("方法调用异常！" + e.getMessage());
		} catch (IllegalArgumentException e) {
			throw new RuntimeException("非法的参数！" + e.getMessage());
		} catch (IllegalAccessException e) {
			throw new RuntimeException("非法的操作！" + e.getMessage());
		} catch (InvocationTargetException e) {
			throw new RuntimeException("调用目标对象异常！" + e.getMessage());
		}

		return excelTemplate;
	}

	/**
	 * 无模板输出对象到文件或流
	 * 
	 * @param objects
	 *            对象列表
	 * @param clazz
	 *            对象类类型
	 * @param isXssf
	 *            是否XSSFWorkbook类型
	 * @return Workbook 返回Workbook
	 */
	private Workbook handlerObject2Excel(List<Object> objects, Class clazz, boolean isXssf) {
		Workbook workbook = null;
		try {
			if (isXssf) {
				// 如果是XSSFWorkbook，则创建XSSFWorkbook类型的Workbook
				workbook = new XSSFWorkbook();
			} else {
				// 如果不是XSSFWorkbook，则创建HSSFWorkbook类型的Workbook
				workbook = new HSSFWorkbook();
			}
			Sheet sheet = workbook.createSheet();
			Row row = sheet.createRow(0);
			// 获取数据列表
			List<ExcelHeader> headers = getHeaderList(clazz);
			// 排序
			Collections.sort(headers);
			// 写标题
			for (int i = 0; i < headers.size(); i++) {
				row.createCell(i).setCellValue(headers.get(i).getTitle());
			}
			// 写数据
			Object object = null;
			for (int i = 0; i < objects.size(); i++) {
				row = sheet.createRow(i + 1);
				object = objects.get(i);
				for (int j = 0; j < headers.size(); j++) {
					row.createCell(j).setCellValue(BeanUtils.getProperty(object, getMethodName(headers.get(j))));
				}
			}
		} catch (IllegalAccessException e) {
			e.printStackTrace();
		} catch (InvocationTargetException e) {
			e.printStackTrace();
		} catch (NoSuchMethodException e) {
			e.printStackTrace();
		}
		return workbook;
	}

	/**
	 * 获取MethodName
	 * 
	 * @param excelHeader
	 * @return
	 */
	private String getMethodName(ExcelHeader excelHeader) {
		// 通过ExcelHeader取得方法名
		String methodName = excelHeader.getMethodName();
		// 将getUsername此种类型的串转换为username
		methodName = methodName.substring(3, methodName.length()).toLowerCase();
		return methodName;
	}

	/**
	 * 获得ExcelHeader的List集合，并按order排序
	 * 
	 * @param clazz
	 * @return
	 */

	private List<ExcelHeader> getHeaderList(Class clazz) {
		List<ExcelHeader> headers = new ArrayList<ExcelHeader>();
		Method[] methods = clazz.getDeclaredMethods();
		for (Method method : methods) {
			String methodName = method.getName();
			if (methodName.startsWith("get")) {
				if (method.isAnnotationPresent(ExcelResources.class)) {
					ExcelResources excelResources = method.getAnnotation(ExcelResources.class);
					headers.add(new ExcelHeader(excelResources.title(), excelResources.order(), methodName));
				}
			}
		}
		return headers;
	}

	/**
	 * 获取某一列所对应的方法名称
	 * 
	 * @param titleRow
	 *            Row
	 * @param clazz
	 *            类类型
	 * @return Map<Integer, String> 列号与方法名的Map
	 */
	private Map<Integer, String> getHeaderMap(Row titleRow, Class clazz) {
		List<ExcelHeader> headers = getHeaderList(clazz);
		Map<Integer, String> maps = new HashMap<Integer, String>();
		for (Cell cell : titleRow) {
			String title = cell.getStringCellValue();
			for (ExcelHeader excelHeader : headers) {
				if (excelHeader.getTitle().trim().equals(title)) {
					maps.put(cell.getColumnIndex(), excelHeader.getMethodName().replaceAll("get", "set"));
					break;
				}
			}
		}
		return maps;
	}

}
