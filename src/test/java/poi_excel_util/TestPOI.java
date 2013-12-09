package poi_excel_util;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;

import org.apache.commons.beanutils.BeanUtils;
import org.junit.Test;

import com.james.model.User;
import com.james.poi_util.ExcelTemplate;
import com.james.poi_util.ExcelUtil;

public class TestPOI {

	@Test
	public void testExcelTemplate() {
		DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd", Locale.CHINA);
		String date1 = dateFormat.format(new Date(System.currentTimeMillis()));
		ExcelTemplate excelTemplate = ExcelTemplate.getInstance().readTemplateByClasspath("/excel/user.xls");
		for (int i = 0; i < 25; i++) {
			excelTemplate.createNewRow();
			excelTemplate.createCell(123.1020001);
			excelTemplate.createCell(true);
			excelTemplate.createCell(date1);
			excelTemplate.createCell(123);
		}
		Map<String, String> datas = new HashMap<String, String>();
		datas.put("title", "测试");
		datas.put("date", date1);
		datas.put("dept", "财务部");
		excelTemplate.replaceConstantData(datas);
		excelTemplate.insertSerialNumber();
		excelTemplate.writeToFile("D:/aaa1.xls");
	}

	/**
	 * 测试exportObject2ExcelByTemplate()
	 */
	@Test
	public void testexportObject2ExcelByTemplate() {
		// 创建User的对象
		List<Object> userList = new ArrayList<Object>();
		for (int i = 1; i < 45; i++) {
			userList.add(new User(i, "james" + i, "james2lee" + i, i + "123456", "m", 34));
		}
		// 获取ExcelUtil实例
		ExcelUtil excelUtil = ExcelUtil.getInstance();
		// 创建title、date、dept等的Map
		DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd", Locale.CHINA);
		String date1 = dateFormat.format(new Date(System.currentTimeMillis()));
		Map<String, String> datas = new HashMap<String, String>();
		datas.put("title", "用户信息列表");
		datas.put("date", date1);
		datas.put("dept", "教务部");
		// 执行从对象导出到Excel文件
		//带格式的Excel文件
		 excelUtil.exportObject2ExcelByTemplate(datas, "/excel/user.xls", "D:/aa.xls", userList, User.class, true);
		 //不带格式的Excel文件
//		excelUtil.exportObject2Excel("D:/a.xls", userList, User.class, false);
	}

	@Test
	public void testRead() {
		//不带格式的Excel文件
//		List<Object> objects = ExcelUtil.getInstance().readExcel2ObjectByPath("D:/aa.xls", User.class);
		//带格式的Excel文件
		List<Object> objects = ExcelUtil.getInstance().readExcel2ObjectByPath("D:/aa.xls", User.class,1,2);
		for (Object object : objects) {
			System.out.println(object);
		}
	}
	
	/**
	 * 测试基于BeanUtils的方式
	 */
	@Test
	public void testBeanUtils() {
		try {
			Class clazz = User.class;
			Object object = clazz.newInstance();
			BeanUtils.copyProperty(object, "username", "james2lee");
			String string = BeanUtils.getProperty(object, "username");
			System.out.println(string);
		} catch (InstantiationException e) {
			e.printStackTrace();
		} catch (IllegalAccessException e) {
			e.printStackTrace();
		} catch (InvocationTargetException e) {
			e.printStackTrace();
		} catch (NoSuchMethodException e) {
			e.printStackTrace();
		}
	}

	/**
	 * 测试基于反射的方式 通过反射来调用对象的方法
	 */
	@Test
	public void testReflection() {
		try {

			Class<User> clazz = User.class;
			Object object = clazz.newInstance();
			// 通过反射调用setUsername方法
			Method method = clazz.getDeclaredMethod("setUsername", String.class);
			method.invoke(object, "李坚勇");
			// 通过反射调用getUsername方法
			Method method2 = clazz.getDeclaredMethod("getUsername");
			System.out.println("method2.invoke(object): " + method2.invoke(object));
			// 通过类型转换来调用getUsername方法
			User user = (User) object;
			System.out.println("user.getUsername(): " + user.getUsername());

		} catch (InstantiationException e) {
			e.printStackTrace();
		} catch (IllegalAccessException e) {
			e.printStackTrace();
		} catch (SecurityException e) {
			e.printStackTrace();
		} catch (NoSuchMethodException e) {
			e.printStackTrace();
		} catch (IllegalArgumentException e) {
			e.printStackTrace();
		} catch (InvocationTargetException e) {
			e.printStackTrace();
		}
	}

	
}
