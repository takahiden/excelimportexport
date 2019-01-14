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

import java.sql.Date;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Time;
import java.sql.Timestamp;
import java.text.ParseException;
import java.text.SimpleDateFormat;

import org.dbunit.dataset.ITable;
import org.dbunit.dataset.datatype.StringDataType;
import org.dbunit.dataset.datatype.TypeCastException;

public class StringTimestampType extends StringDataType {

	public StringTimestampType() {
		super("VARCHAR", 12);
	}

	public static SimpleDateFormat format1 = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss.SSS");
	public static SimpleDateFormat format2 = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss.SSS");
	public static SimpleDateFormat format3 = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
	public static SimpleDateFormat format4 = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");
	public static SimpleDateFormat format5 = new SimpleDateFormat("yyyy/MM/dd HH:mm");
	public static SimpleDateFormat format6 = new SimpleDateFormat("yyyy-MM-dd HH:mm");
	public static SimpleDateFormat format7 = new SimpleDateFormat("yyyy-MM-dd");
	public static SimpleDateFormat format8 = new SimpleDateFormat("yyyy/MM/dd");

	@Override
	public Object getSqlValue(int column, ResultSet resultSet) throws SQLException, TypeCastException {

		Timestamp value = resultSet.getTimestamp(column);
		if ((value == null) || (resultSet.wasNull())) {
			return null;
		}

		return format1.format(value);
	}

	@Override
	public void setSqlValue(Object value, int column, PreparedStatement statement)
			throws SQLException, TypeCastException {

		if ((value == null) || (value == ITable.NO_VALUE)) {
			statement.setString(column, null);
			return;
		}

		if ((value instanceof String)) {
			try {
				statement.setTimestamp(column, new Timestamp(format1.parse((String) value).getTime()));
				return;
			} catch (ParseException e) {
			}
			try {
				statement.setTimestamp(column, new Timestamp(format2.parse((String) value).getTime()));
				return;
			} catch (ParseException e) {
			}
			try {
				statement.setTimestamp(column, new Timestamp(format3.parse((String) value).getTime()));
				return;
			} catch (ParseException e) {
			}
			try {
				statement.setTimestamp(column, new Timestamp(format4.parse((String) value).getTime()));
				return;
			} catch (ParseException e) {
			}
			try {
				statement.setTimestamp(column, new Timestamp(format5.parse((String) value).getTime()));
				return;
			} catch (ParseException e) {
			}
			try {
				statement.setTimestamp(column, new Timestamp(format6.parse((String) value).getTime()));
				return;
			} catch (ParseException e) {
			}
			try {
				statement.setTimestamp(column, new Timestamp(format7.parse((String) value).getTime()));
				return;
			} catch (ParseException e) {
			}
			try {
				statement.setTimestamp(column, new Timestamp(format8.parse((String) value).getTime()));
				return;
			} catch (ParseException e) {
			}
			throw new TypeCastException(value, this);
		}

		if ((value instanceof Date)) {
			statement.setDate(column, (Date) value);
			return;
		}
		if ((value instanceof Time)) {
			statement.setTime(column, (Time) value);
			return;
		}
		if ((value instanceof Timestamp)) {
			statement.setTimestamp(column, (Timestamp) value);
			return;
		}
		if ((value instanceof Long)) {
			Timestamp dateValue = new Timestamp(((Long) value).longValue());
			statement.setTimestamp(column, dateValue);
			return;
		}
		statement.setString(column, asString(value));
	}

}
