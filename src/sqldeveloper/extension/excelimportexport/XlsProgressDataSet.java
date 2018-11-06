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
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import javax.swing.ProgressMonitor;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.dbunit.dataset.DataSetException;
import org.dbunit.dataset.ITable;
import org.dbunit.dataset.ITableIterator;
import org.dbunit.dataset.ITableMetaData;
import org.dbunit.dataset.excel.XlsDataSet;

public class XlsProgressDataSet extends XlsDataSet implements Closeable {

	private ImportDialog owner = null;

	private List<String> tableNames = new ArrayList<String>();
	private int totalRecCount = 0;
	private List<LogBean> logList = null;

	public XlsProgressDataSet(File file, ImportDialog owner, List<LogBean> logList)
			throws IOException, DataSetException {
		super(file);
		this.owner = owner;
		this.logList = logList;

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
				tableNames.add(tableName);
				LogBean logBean = new LogBean();
				logBean.setTableName(tableName);
				logBean.setSql("INSERT...");
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
		pm = new ProgressMonitor(owner, null, null, 0, totalRecCount);
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
			boolean next = iterator.next();
			if (next) {
				if ("(LOG)".equals(iterator.getTable().getTableMetaData().getTableName())) {
					next = iterator.next();
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

			return itable.getValue(paramInt, paramString);
		}

	}

}
