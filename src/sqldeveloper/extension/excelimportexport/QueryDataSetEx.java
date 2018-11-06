package sqldeveloper.extension.excelimportexport;

import java.io.Closeable;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import javax.swing.ProgressMonitor;

import org.dbunit.database.AmbiguousTableNameException;
import org.dbunit.database.IDatabaseConnection;
import org.dbunit.database.QueryDataSet;
import org.dbunit.dataset.DataSetException;
import org.dbunit.dataset.ITable;
import org.dbunit.dataset.ITableIterator;
import org.dbunit.dataset.ITableMetaData;

public class QueryDataSetEx extends QueryDataSet implements Closeable {

	private ExportDialog owner = null;

	private List<LogBean> logList = null;

	public QueryDataSetEx(IDatabaseConnection connection, ExportDialog owner, List<LogBean> logList) {
		super(connection);
		this.owner = owner;
		this.logList = logList;
	}

	private List<String> tableNames = new ArrayList<String>();

	@Override
	public void addTable(String tableName, String query) throws AmbiguousTableNameException {
		tableNames.add(tableName);
		LogBean logBean = new LogBean();
		logBean.setTableName(tableName);
		logBean.setSql(query);
		this.logList.add(logBean);
		super.addTable(tableName, query);
	}

	protected ITableIterator iterator = null;
	ProgressMonitor pm = null;

	@Override
	protected ITableIterator createIterator(boolean reversed) throws DataSetException {
		iterator = super.createIterator(reversed);
		pm = new ProgressMonitor(owner, null, null, 0, tableNames.size());
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

	class ITableIteratorWrapper implements ITableIterator {

		@Override
		public boolean next() throws DataSetException {
			if (logBean != null) {
				logBean.setResult("success");
			}
			return iterator.next();
		}

		@Override
		public ITableMetaData getTableMetaData() throws DataSetException {
			return iterator.getTableMetaData();
		}

		@Override
		public ITable getTable() throws DataSetException {
			itable = iterator.getTable();
			pm.setNote(itable.getTableMetaData().getTableName());
			if (tableNames.size() > 1) {
				pm.setProgress(called++);
			}
			logBean = logList.get(tableIndex++);
			logBean.setTime(new Date());
			return new ITableWrapper();
		}
	}

	public static final String CANCEL_FLG = QueryDataSetEx.class.getName() + "_CANCELED";
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
			if (tableNames.size() == 1) {
				pm.setMaximum(rowCount);
			}
			return rowCount;
		}

		@Override
		public Object getValue(int paramInt, String paramString) throws DataSetException {
			if (pm.isCanceled()) {
				throw new DataSetException(CANCEL_FLG);
			}
			if (tableNames.size() == 1) {
				String currentRecKey = itable.getTableMetaData().getTableName() + ":" + paramInt;
				if (!currentRecKey.equals(preRecKey)) {
					pm.setProgress(++called);
				}
				preRecKey = currentRecKey;
			}
			pm.setNote(itable.getTableMetaData().getTableName() + "( " + (paramInt + 1) + " / " + rowCount + " )");
			return itable.getValue(paramInt, paramString);
		}

	}

	public List<LogBean> getLogList() {
		return logList;
	}

	public void setLogList(List<LogBean> logList) {
		this.logList = logList;
	}
}
