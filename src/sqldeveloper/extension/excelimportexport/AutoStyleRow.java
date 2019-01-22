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

import java.util.Iterator;
import java.util.Spliterator;
import java.util.function.Consumer;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class AutoStyleRow implements Row {

	protected Row delegateRow = null;
	private CellStyle style = null;
	protected AutoFixStyleWorkbook workBook = null;
	protected AutoStyleSheet autoStyleSheet = null;

	public AutoStyleRow(AutoFixStyleWorkbook workBook, AutoStyleSheet autoStyleSheet, Row delegateRow,
			CellStyle style) {
		this.delegateRow = delegateRow;
		this.style = style;
		this.workBook = workBook;
		this.autoStyleSheet = autoStyleSheet;
	}

	public Cell createCell(int column) {
		Cell c = delegateRow.createCell(column);
		c.setCellStyle(style);

		Integer colsize = workBook.colMax.get(autoStyleSheet.getSheetName());
		if (colsize == null || colsize < column) {
			workBook.colMax.put(autoStyleSheet.getSheetName(), column);
		}
		return c;
	}

	public Cell createCell(int column, CellType type) {
		Cell c = delegateRow.createCell(column, type);
		c.setCellStyle(style);

		Integer colsize = workBook.colMax.get(autoStyleSheet.getSheetName());
		if (colsize == null || colsize < column) {
			workBook.colMax.put(autoStyleSheet.getSheetName(), column);
		}
		return c;
	}

	public Iterator<Cell> iterator() {
		return delegateRow.iterator();
	}

	public void forEach(Consumer<? super Cell> action) {
		delegateRow.forEach(action);
	}

	public Spliterator<Cell> spliterator() {
		return delegateRow.spliterator();
	}

	public void removeCell(Cell cell) {
		delegateRow.removeCell(cell);
	}

	public void setRowNum(int rowNum) {
		delegateRow.setRowNum(rowNum);
	}

	public int getRowNum() {
		return delegateRow.getRowNum();
	}

	public Cell getCell(int cellnum) {
		return delegateRow.getCell(cellnum);
	}

	public Cell getCell(int cellnum, MissingCellPolicy policy) {
		return delegateRow.getCell(cellnum, policy);
	}

	public short getFirstCellNum() {
		return delegateRow.getFirstCellNum();
	}

	public short getLastCellNum() {
		return delegateRow.getLastCellNum();
	}

	public int getPhysicalNumberOfCells() {
		return delegateRow.getPhysicalNumberOfCells();
	}

	public void setHeight(short height) {
		delegateRow.setHeight(height);
	}

	public void setZeroHeight(boolean zHeight) {
		delegateRow.setZeroHeight(zHeight);
	}

	public boolean getZeroHeight() {
		return delegateRow.getZeroHeight();
	}

	public void setHeightInPoints(float height) {
		delegateRow.setHeightInPoints(height);
	}

	public short getHeight() {
		return delegateRow.getHeight();
	}

	public float getHeightInPoints() {
		return delegateRow.getHeightInPoints();
	}

	public boolean isFormatted() {
		return delegateRow.isFormatted();
	}

	public CellStyle getRowStyle() {
		return delegateRow.getRowStyle();
	}

	public void setRowStyle(CellStyle style) {
		delegateRow.setRowStyle(style);
	}

	public Iterator<Cell> cellIterator() {
		return delegateRow.cellIterator();
	}

	public Sheet getSheet() {
		return delegateRow.getSheet();
	}

	public int getOutlineLevel() {
		return delegateRow.getOutlineLevel();
	}

	public void shiftCellsRight(int firstShiftColumnIndex, int lastShiftColumnIndex, int step) {
		delegateRow.shiftCellsRight(firstShiftColumnIndex, lastShiftColumnIndex, step);
	}

	public void shiftCellsLeft(int firstShiftColumnIndex, int lastShiftColumnIndex, int step) {
		delegateRow.shiftCellsLeft(firstShiftColumnIndex, lastShiftColumnIndex, step);
	}

}
