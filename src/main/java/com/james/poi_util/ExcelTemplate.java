package com.james.poi_util;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * ExcelTemplate
 * 
 * @author JAMES
 */
public class ExcelTemplate {
	/**
	 * 数据开始的行
	 */
	public final static String DATA_LINE = "datas";
	/**
	 * 默认样式
	 */
	public final static String DEFAULT_STYLE = "defaultstyles";
	/**
	 * 样式
	 */
	public final static String STYLE = "styles";
	/**
	 * 序号
	 */
	public final static String SERIALNUMBER = "SN";

	private Workbook workbook;

	private Sheet sheet;

	/**
	 * 初始化的列下标，从0开始
	 */
	private Integer initColIndex;
	/**
	 * 初始化的行下标，从0开始
	 */
	private Integer initRowIndex;
	/**
	 * 当前列下标
	 */
	private Integer curColIndex;
	/**
	 * 当前行下标
	 */
	private Integer curRowIndex;
	/**
	 * 当前行对象
	 */
	private Row curRow;
	/**
	 * 最后一行的数据
	 */
	private Integer lastRowIndex;

	/**
	 * 默认样式
	 */
	private CellStyle defaultStyle;

	/**
	 * 默认行高
	 */
	private float rowHeight;

	/**
	 * 储存某一列所对应的样式Map<列号,样式>
	 */
	private Map<Integer, CellStyle> styles;

	/**
	 * 序号的列
	 */
	private Integer serialNumberColIndex;

	

	// ======================================================================================

	/**
	 * 单例
	 */
	private static ExcelTemplate excelTemplate = null;

	/**
	 * 默认无参构造方法
	 */
	private ExcelTemplate() {}

	/**
	 * 提供ExcelTemplate的实例对象
	 * 
	 * @return
	 */
	public static ExcelTemplate getInstance() {
		if (excelTemplate == null) {
			excelTemplate = new ExcelTemplate();
		}
		return excelTemplate;
	}

	/**
	 * 读取Classpath下的模板文件
	 * 
	 * @param classpath
	 */
	public ExcelTemplate readTemplateByClasspath(String classpath) {
		try {
			workbook = WorkbookFactory.create(ExcelTemplate.class.getResourceAsStream(classpath));
			initTemplate();
		} catch (InvalidFormatException e) {
			throw new RuntimeException("模板文件格式错误，请检查！");
		} catch (IOException e) {
			throw new RuntimeException("模板文件不存在，请检查！");
		}
		return excelTemplate;
	}

	/**
	 * 读取path指定的模板文件
	 * 
	 * @param path
	 */
	public ExcelTemplate readTemplateByPath(String path) {
		try {
			workbook = WorkbookFactory.create(new File(path));
			initTemplate();
		} catch (InvalidFormatException e) {
			throw new RuntimeException("模板文件格式错误，请检查！");
		} catch (IOException e) {
			throw new RuntimeException("模板文件不存在，请检查！");
		}
		return excelTemplate;
	}

	/**
	 * 数据写入到文件
	 * 
	 * @param filepath
	 */
	public void writeToFile(String filepath) {
		FileOutputStream fileOutputStream = null;
		try {
			fileOutputStream = new FileOutputStream(new File(filepath));
			workbook.write(fileOutputStream);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
			throw new RuntimeException(filepath + "，文件没找到！" + e.getMessage());
		} catch (IOException e) {
			e.printStackTrace();
			throw new RuntimeException("OutputStream，IO异常！" + e.getMessage());
		}
		finally {
			try {
				if (fileOutputStream != null) {
					fileOutputStream.close();
				}
			} catch (IOException e) {
				e.printStackTrace();
				throw new RuntimeException("OutputStream关闭流时，IO异常！" + e.getMessage());
			}
		}

	}

	/**
	 * 数据写入到输出流
	 * 
	 * @param outputStream
	 */
	public void writeToStream(OutputStream outputStream) {
		try {
			workbook.write(outputStream);
		} catch (IOException e) {
			e.printStackTrace();
			throw new RuntimeException("OutputStream输出流，IO异常！");
		}
	}

	/**
	 * 创建新列
	 * @param value
	 */
	public void createCell(String value) {
		Cell cell = curRow.createCell(curColIndex);
		setCellStyle(cell);
		cell.setCellValue(value);
		curColIndex++;
	}
	
	/**
	 * 创建新列
	 * @param value
	 */
	public void createCell(Object value) {
		Cell cell = curRow.createCell(curColIndex);
		setCellStyle(cell);
		cell.setCellValue(value.toString());
		curColIndex++;
	}
	
	/**
	 * 创建新列
	 * @param value
	 */
	public void createCell(RichTextString value) {
		Cell cell = curRow.createCell(curColIndex);
		setCellStyle(cell);
		cell.setCellValue(value);
		curColIndex++;
	}
	
	/**
	 * 创建新列
	 * @param value
	 */
	public void createCell(Boolean value) {
		Cell cell = curRow.createCell(curColIndex);
		setCellStyle(cell);
		cell.setCellValue(value);
		curColIndex++;
	}
	
	/**
	 * 创建新列
	 * @param value
	 */
	public void createCell(Integer value) {
		Cell cell = curRow.createCell(curColIndex);
		setCellStyle(cell);
		cell.setCellValue(value);
		curColIndex++;
	}
	/**
	 * 创建新列
	 * @param value
	 */
	public void createCell(Date value) {
		Cell cell = curRow.createCell(curColIndex);
		setCellStyle(cell);
		cell.setCellValue(value);
		curColIndex++;
	}
	
	/**
	 * 创建新列
	 * @param value
	 */
	public void createCell(Double value) {
		Cell cell = curRow.createCell(curColIndex);
		setCellStyle(cell);
		cell.setCellValue(value);
		curColIndex++;
	}
	
	/**
	 * 创建新列
	 * @param value
	 */
	public void createCell(Calendar value) {
		Cell cell = curRow.createCell(curColIndex);
		setCellStyle(cell);
		cell.setCellValue(value);
		curColIndex++;
	}
	
	
	/**
	 * 设置CellStyle
	 @param cell
	 */
	private void setCellStyle(Cell cell) {
		// 如果Map里面包含有当前列号，则使用当前列的专用样式
		if (styles.containsKey(curColIndex)) {
			cell.setCellStyle(styles.get(curColIndex));
			// 否则使用样式为默认样式
		} else {
			cell.setCellStyle(defaultStyle);
		}
	}
	
	/**
	 * 创建新行
	 */
	public void createNewRow() {
		if (lastRowIndex > curRowIndex && curRowIndex != initRowIndex) {
			/**
			 * sheet.shiftRows(从哪一行开始, 到哪一行结束, 插入多少行, 高度是否与上一行一致, 宽度是否与上一列一致);
			 */
			sheet.shiftRows(curRowIndex, lastRowIndex, 1, true, true);
			lastRowIndex++;
		}
		curRow = sheet.createRow(curRowIndex);
		curRow.setHeightInPoints(rowHeight);
		curRowIndex++;
		curColIndex = initColIndex;
	}
	
	

	/**
	 * 插入序号
	 */
	public void insertSerialNumber() {
		int index = 1;
		Row row = null;
		Cell cell = null;
		for (int i = initRowIndex; i < curRowIndex; i++) {
			row = sheet.getRow(i);
			cell = row.createCell(serialNumberColIndex);
			//设置样式
			setCellStyle(cell);
			//设置序号
			cell.setCellValue(index++);
		}
	}

	/**
	 * 根据Map<String,String>替换相应的常量
	 */
	public void replaceConstantData(Map<String, String> datas) {
		if (datas == null) {
			return;
		}
		for (Row row : sheet) {
			for (Cell cell : row) {
				if (cell.getCellType() != Cell.CELL_TYPE_STRING) {
					continue;
				}
				String string = cell.getStringCellValue().trim();
				if (string.startsWith("#") && string.endsWith("#")) {
					if (datas.containsKey(string.substring(1, string.length() - 1))) {
						cell.setCellValue(datas.get(string.substring(1, string.length() - 1)));
					}
				}
			}
		}
	}

	/**
	 * 初始化模板文件
	 */
	private void initTemplate() {
		sheet = workbook.getSheetAt(0);
		initConfigData();
		lastRowIndex = sheet.getLastRowNum();
		curRow = sheet.createRow(curRowIndex);
	}

	/**
	 * 初始化配置数据
	 */
	private void initConfigData() {
		boolean findData = false;
		boolean findSerialNumber = false;
		for (Row row : sheet) {
			if (findData) {
				break;
			}
			for (Cell cell : row) {
				if (cell.getCellType() != Cell.CELL_TYPE_STRING) {
					continue;
				}
				String string = cell.getStringCellValue().trim();
				// 是否有找到SN
				if (string.equals(SERIALNUMBER)) {
					serialNumberColIndex = cell.getColumnIndex();
					cell.getRowIndex();
					findSerialNumber = true;
				}
				if (string.equals(DATA_LINE)) {
					initColIndex = cell.getColumnIndex();
					initRowIndex = row.getRowNum();
					curColIndex = initColIndex;
					curRowIndex = initRowIndex;
					findData = true;
					defaultStyle = cell.getCellStyle();
					rowHeight = row.getHeightInPoints();
					initStyles();
					break;
				}
			}
		}
		// 如果没有找到模板文件中的SN
		if (!findSerialNumber) {
			initSerialNumber();
		}
	}

	/**
	 * 初始化序列号
	 */
	private void initSerialNumber() {
		for (Row row : sheet) {
			for (Cell cell : row) {
				if (cell.getCellType() != Cell.CELL_TYPE_STRING) {
					continue;
				}
				String string = cell.getStringCellValue().trim();
				// 是否有找到SN
				if (string.equals(SERIALNUMBER)) {
					serialNumberColIndex = cell.getColumnIndex();
				}
			}
		}
	}

	/**
	 * 初始化每一列的样式
	 */
	private void initStyles() {
		styles = new HashMap<Integer, CellStyle>();
		for (Row row : sheet) {
			for (Cell cell : row) {
				if (cell.getCellType() != Cell.CELL_TYPE_STRING) {
					continue;
				}
				String string = cell.getStringCellValue().trim();
				// 如果取得的内容为“DEFAULT_STYLE”，则当前行的样式使用默认样式
				if (string.equals(DEFAULT_STYLE)) {
					defaultStyle = cell.getCellStyle();
				}
				// 如果取得的内容为“STYLE”，则当前行的样式使用STYLE样式
				if (string.equals(STYLE)) {
					styles.put(cell.getColumnIndex(), cell.getCellStyle());
				}
			}
		}
	}
}
