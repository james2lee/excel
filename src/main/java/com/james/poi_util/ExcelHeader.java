package com.james.poi_util;

import java.io.Serializable;

/**
 * 将title、order、methodName封装到ExcelHeader类中，并按order排序
  @author JAMES
 */
public class ExcelHeader implements Serializable, Comparable<ExcelHeader> {
	private static final long serialVersionUID = 1L;

	private String title;
	private Integer order;
	private String methodName;

	public ExcelHeader() {
		super();

	}

	public ExcelHeader(String title, Integer order, String methodName) {
		super();
		this.title = title;
		this.order = order;
		this.methodName = methodName;
	}

	public String getTitle() {
		return title;
	}

	public void setTitle(String title) {
		this.title = title;
	}

	public Integer getOrder() {
		return order;
	}

	public void setOrder(Integer order) {
		this.order = order;
	}

	public String getMethodName() {
		return methodName;
	}

	public void setMethodName(String methodName) {
		this.methodName = methodName;
	}

	public int compareTo(ExcelHeader excelHeader) {
		return order > excelHeader.order ? 1 : (order < excelHeader.order ? -1 : 0);
	}

	@Override
	public String toString() {
		return "ExcelHeader [title=" + title + ", order=" + order + ", methodName=" + methodName + "]";
	}

}
