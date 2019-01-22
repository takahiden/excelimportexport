/*
Copyright (c) 2018, Takahiden. All rights reserved. 

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License. 
*/
package sqldeveloper.extension.excelimportexport;

public class LogBean {

	public String getTime() {
		return time;
	}

	public void setTime(String time) {
		this.time = time;
	}

	public String getTableName() {
		return tableName;
	}

	public void setTableName(String tableName) {
		this.tableName = tableName;
	}

	public String getResult() {
		return result;
	}

	public void setResult(String result) {
		this.result = result;
	}

	private String time = null;
	private String tableName = null;

	public String getSql() {
		return sql;
	}

	public void setSql(String sql) {
		this.sql = sql;
	}

	private String sql = null;
	private String result = null;

	public String getElse1() {
		return else1;
	}

	public void setElse1(String else1) {
		this.else1 = else1;
	}

	public String getElse2() {
		return else2;
	}

	public void setElse2(String else2) {
		this.else2 = else2;
	}

	public String getElse3() {
		return else3;
	}

	public void setElse3(String else3) {
		this.else3 = else3;
	}

	public String getElse4() {
		return else4;
	}

	public void setElse4(String else4) {
		this.else4 = else4;
	}

	public String getElse5() {
		return else5;
	}

	public void setElse5(String else5) {
		this.else5 = else5;
	}

	private String else1 = null;
	private String else2 = null;
	private String else3 = null;
	private String else4 = null;
	private String else5 = null;
}
