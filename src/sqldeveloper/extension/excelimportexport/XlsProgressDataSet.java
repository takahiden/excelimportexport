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

import java.io.Closeable;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

import javax.swing.ProgressMonitor;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.dbunit.dataset.Column;
import org.dbunit.dataset.DataSetException;
import org.dbunit.dataset.ITable;
import org.dbunit.dataset.ITableIterator;
import org.dbunit.dataset.ITableMetaData;
import org.dbunit.dataset.excel.XlsDataSet;

import oracle.dbtools.raptor.utils.Connections;
import oracle.ide.log.LogManager;
import oracle.javatools.db.DBException;

public class XlsProgressDataSet extends XlsDataSet implements Closeable {

	private ImportDialog owner = null;

	private int totalRecCount = 0;
	private List<LogBean> logList = null;
	private boolean deleteBeforeImport;

	public XlsProgressDataSet(File file, ImportDialog owner, List<LogBean> logList, boolean deleteBeforeImport)
			throws IOException, DataSetException {
		super(file);
		this.owner = owner;
		this.logList = logList;
		this.deleteBeforeImport = deleteBeforeImport;
		Workbook workbook;
		try {
			workbook = WorkbookFactory.create(new FileInputStream(file));
		} catch (InvalidFormatException e) {
			throw new IOException(e);
		}
		int sheetCount = workbook.getNumberOfSheets();
		for (int i = 0; i < sheetCount; i++) {
			String tableName = workbook.getSheetName(i);
			if (!"(LOG)".equals(tableName)) {
				LogBean logBean = new LogBean();
				logBean.setTableName(tableName);
				if (deleteBeforeImport) {
					logBean.setSql("DELETE and INSERT...");
				} else {
					logBean.setSql("INSERT...");
				}
				this.logList.add(logBean);
				totalRecCount += workbook.getSheetAt(i).getLastRowNum();
			}
		}

	}

	protected ITableIterator iterator = null;
	ProgressMonitor pm = null;

	@Override
	protected ITableIterator createIterator(boolean reversed) throws DataSetException {
		iterator = super.createIterator(reversed);
		if (pm == null) {
			pm = new ProgressMonitor(owner, null, null, 0, totalRecCount);
		}
		pm.setMillisToDecideToPopup(1);
		pm.setMillisToPopup(1);
		return new ITableIteratorWrapper();
	}

	@Override
	public void close() throws IOException {
		if (pm != null) {
			pm.close();
		}
	}

	private int called = 0;
	private ITable itable = null;

	private int tableIndex = 0;
	private LogBean logBean = null;

	private int calledIterator = 0;

	class ITableIteratorWrapper implements ITableIterator {

		ITableIteratorWrapper() {
			calledIterator++;
			tableIndex = 0;
		}

		@Override
		public boolean next() throws DataSetException {
			if (logBean != null) {
				logBean.setResult("success");
			}
			boolean next = iterator.next();
			if (next) {
				String tableName = iterator.getTable().getTableMetaData().getTableName();
				if ("(LOG)".equals(tableName)) {
					next = next();
				} else if (tableName != null && tableName.startsWith("(SQL")) {

					// 削除の場合は2回イテレータが呼ばれ、2回目で呼び出すための回避
					if ((deleteBeforeImport && calledIterator > 1) || !deleteBeforeImport) {

						// SQLシートは個別実行
						itable = iterator.getTable();
						pm.setNote(itable.getTableMetaData().getTableName());
						rowCount = itable.getRowCount();

						for (int row = 0; row < rowCount; row++) {
							Column[] columns = itable.getTableMetaData().getColumns();

							for (int i = 0; i < columns.length; i++) {
								Object query = getExecuteQuery(row, columns[i].getColumnName());
								if (query != null && query instanceof String) {
									String[] qs = ((String) query).split(";");
									for (String q : qs) {
										if (q.trim().length() > 0) {
											executeQuery(q);
										}
									}
								}
							}
						}
						logBean = logList.get(tableIndex);
						logBean.setSql("EXECUTE SQL...");
						logBean.setTime(new Date());

						tableIndex++;
					}
					next = next();
				}
			}
			return next;
		}

		@Override
		public ITableMetaData getTableMetaData() throws DataSetException {
			return iterator.getTableMetaData();
		}

		@Override
		public ITable getTable() throws DataSetException {
			itable = iterator.getTable();
			pm.setNote(itable.getTableMetaData().getTableName());
			rowCount = itable.getRowCount();
			logBean = logList.get(tableIndex++);
			logBean.setTime(new Date());
			return new ITableWrapper();
		}
	}

	public static final String CANCEL_FLG = XlsProgressDataSet.class.getName() + "_CANCELED";
	private int rowCount = 0;
	String preRecKey = null;

	class ITableWrapper implements ITable {

		@Override
		public ITableMetaData getTableMetaData() {
			return itable.getTableMetaData();
		}

		@Override
		public int getRowCount() {
			rowCount = itable.getRowCount();
			return rowCount;
		}

		@Override
		public Object getValue(int paramInt, String paramString) throws DataSetException {
			if (pm.isCanceled()) {
				throw new DataSetException(CANCEL_FLG);
			}
			String currentRecKey = itable.getTableMetaData().getTableName() + ":" + paramInt;
			if (!currentRecKey.equals(preRecKey)) {
				pm.setProgress(++called);
			}
			pm.setNote(itable.getTableMetaData().getTableName() + "( " + (paramInt + 1) + " / " + rowCount + " )");
			preRecKey = currentRecKey;

			try {
				return itable.getValue(paramInt, paramString);
			} catch (NullPointerException e) {
				return null;
			}
		}

	}

	public Object getExecuteQuery(int paramInt, String paramString) throws DataSetException {
		if (pm.isCanceled()) {
			throw new DataSetException(CANCEL_FLG);
		}
		String currentRecKey = itable.getTableMetaData().getTableName() + ":" + paramInt;
		if (!currentRecKey.equals(preRecKey)) {
			pm.setProgress(++called);
		}
		pm.setNote(itable.getTableMetaData().getTableName() + "( " + (paramInt + 1) + " / " + rowCount + " )");
		preRecKey = currentRecKey;

		return itable.getValue(paramInt, paramString);
	}

	public Integer executeQuery(String query) throws DataSetException {
		Statement st = null;
		int result = 0;

		try {
			Connection conn = getConnection();
			st = conn.createStatement();
			query = query.trim();
			result = st.executeUpdate(query);

		} catch (DBException | SQLException e) {
			LogMessage("WARN", e.getMessage());
			throw new DataSetException(e);
		} finally {
			if (st != null) {
				try {
					st.close();
				} catch (SQLException e) {
					;
				}
			}
		}
		return result;
	}

	public Connection getConnection() throws DBException {
		String connectionName = owner.connectionName;
		return Connections.getInstance().getConnection(connectionName);
	}

	private static final void LogMessage(String level, String msg) {
		DateFormat dateFormat = new SimpleDateFormat("HH:mm:ss");
		Date date = new Date();
		String currentTime = "[" + dateFormat.format(date) + "] ";
		LogManager.getLogManager().getMsgPage().log("EXPORT " + currentTime + level + ": " + msg + "\n");
	}
}
