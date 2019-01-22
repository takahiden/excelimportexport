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

import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.text.DecimalFormatSymbols;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.NoSuchElementException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.dbunit.dataset.AbstractTable;
import org.dbunit.dataset.Column;
import org.dbunit.dataset.DataSetException;
import org.dbunit.dataset.DefaultTableMetaData;
import org.dbunit.dataset.ITableMetaData;
import org.dbunit.dataset.datatype.DataType;
import org.dbunit.dataset.datatype.DataTypeException;

import com.monitorjbl.xlsx.impl.StreamingSheet;

class XlsStreamTable extends AbstractTable {

	protected ITableMetaData metaData = null;
	protected Sheet sheet;
	protected DecimalFormatSymbols symbols = new DecimalFormatSymbols();

	protected int rowCount = 0;
	protected int currentRow = -1;
	protected List<Object> currentColumns = new ArrayList<Object>();

	protected Iterator<Row> rowIterator = null;

	public XlsStreamTable(String sheetName, Sheet sheet) throws DataSetException {
		this.sheet = sheet;
		this.rowCount = sheet.getLastRowNum();
		this.symbols.setDecimalSeparator('.');

		if (this.rowCount >= 0) {
			this.rowIterator = ((StreamingSheet) sheet).rowIterator();
			Row r = loadCurrentRow();
			this.metaData = createMetaData(sheetName, r);
		}
		if (this.metaData == null) {
			this.metaData = new DefaultTableMetaData(sheetName, new Column[0]);
		}
	}

	private static ITableMetaData createMetaData(String tableName, Row sampleRow) {

		List<Column> columnList = new ArrayList<Column>();
		for (Cell cell : sampleRow) {
			if (cell == null) {
				break;
			}

			String columnName = cell.getRichStringCellValue().getString();
			if (columnName != null) {
				columnName = columnName.trim();
			}

			if (columnName.length() <= 0) {
				break;
			}

			Column column = new Column(columnName, DataType.UNKNOWN);
			columnList.add(column);
		}
		Column[] columns = (Column[]) columnList.toArray(new Column[0]);
		return new DefaultTableMetaData(tableName, columns);
	}

	public int getRowCount() {
		return this.rowCount;
	}

	public ITableMetaData getTableMetaData() {
		return this.metaData;
	}

	protected Row loadCurrentRow() throws DataSetException {
		if (!this.rowIterator.hasNext()) {
			throw new DataSetException(new NoSuchElementException(
					"Error at sheet=" + this.sheet.getSheetName() + ", row=" + this.currentRow));
		}
		Row row = this.rowIterator.next();
		this.currentColumns.clear();
		int index = 0;
		int cols = 0;
		if ("(LOG)".equals(this.sheet.getSheetName())) {
			cols = 9;
		} else if (row.getRowNum() == 0) {
			cols = row.getLastCellNum();
		} else {
			cols = getTableMetaData().getColumns().length;
		}

		for (int i = 0; i < cols; i++) {
			this.currentColumns.add(null);
		}

		for (Cell cell : row) {
			if (index >= cols) {
				break;
			}
			Object value = getCellValue(cell);
			this.currentColumns.set(cell.getColumnIndex(), value);
			index++;
		}
		return row;
	}

	protected Object getCellValue(Cell cell) throws DataSetException {
		if (cell == null) {
			return null;
		}

		CellType type = cell.getCellType();
		switch (type) {
		case NUMERIC:
			CellStyle style = cell.getCellStyle();
			if (DateUtil.isCellDateFormatted(cell)) {
				return getDateValue(cell);
			}
			if ("####################".equals(style.getDataFormatString())) {

				return getDateValueFromJavaNumber(cell);
			}

			return getNumericValue(cell);

		case STRING:
			return cell.getRichStringCellValue().getString();

		case FORMULA:
			throw new DataTypeException("Formula not supported at sheet=" + this.sheet.getSheetName() + ", row="
					+ cell.getRowIndex() + ", column=" + cell.getColumnIndex());

		case BLANK:
			return null;

		case _NONE:
			return null;

		case BOOLEAN:
			return cell.getBooleanCellValue() ? Boolean.TRUE : Boolean.FALSE;

		default:
			throw new DataTypeException("Error at sheet=" + this.sheet.getSheetName() + ", row=" + cell.getRowIndex()
					+ ", column=" + cell.getColumnIndex());
		}
	}

	public Object getValue(int row, String column) throws DataSetException {

		assertValidRowIndex(row);

		if (this.currentRow == row) {
			int columnIndex = getColumnIndex(column);
			return this.currentColumns.get(columnIndex);

		} else if ((this.currentRow + 1) == row) {
			this.currentRow = this.currentRow + 1;
			loadCurrentRow();
			int columnIndex = getColumnIndex(column);
			return this.currentColumns.get(columnIndex);

		}
		throw new DataTypeException("Current read index not match getValue at sheet=" + this.sheet.getSheetName()
				+ ", row=" + row + ", column=" + column);
	}

	public Object getValue(int row, int columnIndex) throws DataSetException {

		assertValidRowIndex(row);

		if (this.currentRow == row) {
			return this.currentColumns.get(columnIndex);

		} else if ((this.currentRow + 1) == row) {
			this.currentRow = this.currentRow + 1;
			loadCurrentRow();
			return this.currentColumns.get(columnIndex);

		}
		throw new DataTypeException("Current read index not match getValue at sheet=" + this.sheet.getSheetName()
				+ ", row=" + row + ", columnIndex=" + columnIndex);
	}

	protected Object getDateValueFromJavaNumber(Cell cell) {

		double numericValue = cell.getNumericCellValue();
		BigDecimal numericValueBd = new BigDecimal(String.valueOf(numericValue));
		numericValueBd = stripTrailingZeros(numericValueBd);
		return new Long(numericValueBd.longValue());
	}

	protected Object getDateValue(Cell cell) {

		double numericValue = cell.getNumericCellValue();
		Date date = DateUtil.getJavaDate(numericValue);
		return new Long(date.getTime());
	}

	private BigDecimal stripTrailingZeros(BigDecimal value) {
		if (value.scale() <= 0) {
			return value;
		}

		String valueAsString = String.valueOf(value);
		int idx = valueAsString.indexOf(".");
		if (idx == -1) {
			return value;
		}

		for (int i = valueAsString.length() - 1; i > idx; i--) {
			if (valueAsString.charAt(i) == '0') {
				valueAsString = valueAsString.substring(0, i);
			} else {
				if (valueAsString.charAt(i) != '.')
					break;
				valueAsString = valueAsString.substring(0, i);

				break;
			}
		}

		BigDecimal result = new BigDecimal(valueAsString);
		return result;
	}

	protected BigDecimal getNumericValue(Cell cell) {

		String formatString = cell.getCellStyle().getDataFormatString();
		String resultString = null;
		double cellValue = cell.getNumericCellValue();

		if (formatString != null) {
			if ((!formatString.equals("General")) && (!formatString.equals("@"))) {

				DecimalFormat nf = new DecimalFormat(formatString, this.symbols);
				resultString = nf.format(cellValue);
			}
		}

		BigDecimal result;
		if (resultString != null) {

			try {
				result = new BigDecimal(resultString);
			} catch (NumberFormatException e) {

				result = toBigDecimal(cellValue);
			}
		} else {
			result = toBigDecimal(cellValue);
		}
		return result;
	}

	private BigDecimal toBigDecimal(double cellValue) {
		String resultString = String.valueOf(cellValue);

		if (resultString.endsWith(".0")) {
			resultString = resultString.substring(0, resultString.length() - 2);
		}
		BigDecimal result = new BigDecimal(resultString);
		return result;
	}

	public String toString() {
		StringBuilder sb = new StringBuilder();
		sb.append(getClass().getName()).append("[");
		sb.append("_metaData=").append(this.metaData == null ? "null" : this.metaData.toString());
		sb.append(", _sheet=").append("" + this.sheet);
		sb.append(", symbols=").append("" + this.symbols);
		sb.append("]");
		return sb.toString();
	}
}
