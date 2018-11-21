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

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import oracle.dbtools.raptor.utils.Connections;
import oracle.ide.log.LogManager;
import oracle.javatools.db.DBException;

public class PKListUtil {

	public static String getPKColumnOrderBy(String connectionName, String tableName) {
		StringBuilder retStr = new StringBuilder("");
		List<String> columns = getPKColumnNames(connectionName, tableName);

		for (int i = 0; i < columns.size(); i++) {
			if (i == 0) {
				retStr.append(" ORDER BY ");
				retStr.append(columns.get(i));
			} else {
				retStr.append(" ,");
				retStr.append(columns.get(i));
			}
		}
		return retStr.toString();
	}

	public static List<String> getPKColumnNames(String connectionName, String tableName) {
		PreparedStatement ps = null;
		ResultSet rset = null;
		List<String> retList = new ArrayList<String>();
		try {
			Connection conn = getConnection(connectionName);

			String sql = "SELECT COLUMN_NAME FROM USER_CONS_COLUMNS WHERE TABLE_NAME = ? AND CONSTRAINT_NAME IN ( SELECT CONSTRAINT_NAME FROM USER_CONSTRAINTS WHERE TABLE_NAME = ? AND CONSTRAINT_TYPE = 'P') ORDER BY POSITION";
			ps = conn.prepareStatement(sql);

			ps.setString(1, tableName);
			ps.setString(2, tableName);
			rset = ps.executeQuery();

			while (rset.next()) {
				retList.add(rset.getString(1));
			}
		} catch (DBException | SQLException e) {
			LogMessage("WARN", e.getMessage());
		} finally {
			if (rset != null) {
				try {
					rset.close();
				} catch (SQLException e) {
					;
				}
			}
			if (ps != null) {
				try {
					ps.close();
				} catch (SQLException e) {
					;
				}
			}
		}
		return retList;
	}

	public static Connection getConnection(String connectionName) throws DBException {
		return Connections.getInstance().getConnection(connectionName);
	}

	private static final void LogMessage(String level, String msg) {
		DateFormat dateFormat = new SimpleDateFormat("HH:mm:ss");
		Date date = new Date();
		String currentTime = "[" + dateFormat.format(date) + "] ";
		LogManager.getLogManager().getMsgPage().log("EXPORT " + currentTime + level + ": " + msg + "\n");
	}
}
