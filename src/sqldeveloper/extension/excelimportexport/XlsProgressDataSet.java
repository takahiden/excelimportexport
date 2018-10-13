package sqldeveloper.extension.excelimportexport;

import java.io.Closeable;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
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

	public XlsProgressDataSet(File file, ImportDialog owner) throws IOException, DataSetException {
		super(file);
		this.owner = owner;
		Workbook workbook;
		try {
			workbook = WorkbookFactory.create(new FileInputStream(file));
		} catch (InvalidFormatException e) {
			throw new IOException(e);
		}
		int sheetCount = workbook.getNumberOfSheets();
		for (int i = 0; i < sheetCount; i++) {
			tableNames.add(workbook.getSheetName(i));
			totalRecCount += workbook.getSheetAt(i).getLastRowNum();
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

	class ITableIteratorWrapper implements ITableIterator {

		@Override
		public boolean next() throws DataSetException {
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
			rowCount = itable.getRowCount();
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
