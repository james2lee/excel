package com.james.model;

import java.io.Serializable;

import com.james.poi_util.ExcelResources;

public class User implements Serializable {

	private static final long serialVersionUID = 1L;

	private Integer id;
	private String username;
	private String nickname;
	private String password;
	private String gender;
	private Integer age;

	public User() {
		super();

	}

	public User(Integer id, String username, String nickname, String password, String gender, Integer age) {
		super();
		this.id = id;
		this.username = username;
		this.nickname = nickname;
		this.password = password;
		this.gender = gender;
		this.age = age;
	}

	@ExcelResources(title = "用户标识", order = 1)
	public Integer getId() {
		return id;
	}

	public void setId(Integer id) {
		this.id = id;
	}

	@ExcelResources(title = "用户名称", order = 3)
	public String getUsername() {
		return username;
	}

	public void setUsername(String username) {
		this.username = username;
	}

	@ExcelResources(title = "用户昵称", order = 2)
	public String getNickname() {
		return nickname;
	}

	public void setNickname(String nickname) {
		this.nickname = nickname;
	}

	@ExcelResources(title = "用户密码", order = 4)
	public String getPassword() {
		return password;
	}

	public void setPassword(String password) {
		this.password = password;
	}

	@ExcelResources(title = "性别", order = 5)
	public String getGender() {
		return gender;
	}

	public void setGender(String gender) {
		this.gender = gender;
	}

	@ExcelResources(title = "年龄")
	public Integer getAge() {
		return age;
	}

	public void setAge(Integer age) {
		this.age = age;
	}

	@Override
	public String toString() {
		return "User [id=" + id + ", username=" + username + ", nickname=" + nickname + ", password=" + password + ", gender=" + gender + ", age=" + age + "]";
	}

}
