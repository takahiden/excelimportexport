/**
 * 
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

	public static SimpleDateFormat formatNormal = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss.SSS");
	public static SimpleDateFormat format2 = new SimpleDateFormat("yyyy-MM-dd");
	public static SimpleDateFormat format3 = new SimpleDateFormat("yyyy/MM/dd");
	public static SimpleDateFormat format4 = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");
	public static SimpleDateFormat format5 = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss.SSS");
	public static SimpleDateFormat format6 = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss.SSS");
	
	@Override
	public Object getSqlValue(int column, ResultSet resultSet) throws SQLException, TypeCastException {

		Timestamp value = resultSet.getTimestamp(column);
		if ((value == null) || (resultSet.wasNull())) {
			return null;
		}

		return formatNormal.format(value);
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
				statement.setTimestamp(column, new Timestamp(formatNormal.parse((String) value).getTime()));
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
		statement.setString(column, asString(value));
	}

}
