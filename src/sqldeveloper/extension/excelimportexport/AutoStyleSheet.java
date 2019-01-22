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

import java.util.Collection;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Spliterator;
import java.util.function.Consumer;

import org.apache.poi.ss.usermodel.AutoFilter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellRange;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Footer;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.SheetConditionalFormatting;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.PaneInformation;
import org.apache.poi.xssf.streaming.SXSSFSheet;

public class AutoStyleSheet implements Sheet {

	public Iterator<Row> iterator() {
		return delegateSheet.iterator();
	}

	public void forEach(Consumer<? super Row> action) {
		delegateSheet.forEach(action);
	}

	public Spliterator<Row> spliterator() {
		return delegateSheet.spliterator();
	}

	public void removeRow(Row row) {
		delegateSheet.removeRow(row);
	}

	public Row getRow(int rownum) {
		return delegateSheet.getRow(rownum);
	}

	public int getPhysicalNumberOfRows() {
		return delegateSheet.getPhysicalNumberOfRows();
	}

	public int getFirstRowNum() {
		return delegateSheet.getFirstRowNum();
	}

	public int getLastRowNum() {
		return delegateSheet.getLastRowNum();
	}

	public void setColumnHidden(int columnIndex, boolean hidden) {
		delegateSheet.setColumnHidden(columnIndex, hidden);
	}

	public boolean isColumnHidden(int columnIndex) {
		return delegateSheet.isColumnHidden(columnIndex);
	}

	public void setRightToLeft(boolean value) {
		delegateSheet.setRightToLeft(value);
	}

	public boolean isRightToLeft() {
		return delegateSheet.isRightToLeft();
	}

	public void setColumnWidth(int columnIndex, int width) {
		delegateSheet.setColumnWidth(columnIndex, width);
	}

	public int getColumnWidth(int columnIndex) {
		return delegateSheet.getColumnWidth(columnIndex);
	}

	public float getColumnWidthInPixels(int columnIndex) {
		return delegateSheet.getColumnWidthInPixels(columnIndex);
	}

	public void setDefaultColumnWidth(int width) {
		delegateSheet.setDefaultColumnWidth(width);
	}

	public int getDefaultColumnWidth() {
		return delegateSheet.getDefaultColumnWidth();
	}

	public short getDefaultRowHeight() {
		return delegateSheet.getDefaultRowHeight();
	}

	public float getDefaultRowHeightInPoints() {
		return delegateSheet.getDefaultRowHeightInPoints();
	}

	public void setDefaultRowHeight(short height) {
		delegateSheet.setDefaultRowHeight(height);
	}

	public void setDefaultRowHeightInPoints(float height) {
		delegateSheet.setDefaultRowHeightInPoints(height);
	}

	public CellStyle getColumnStyle(int column) {
		return delegateSheet.getColumnStyle(column);
	}

	public int addMergedRegion(CellRangeAddress region) {
		return delegateSheet.addMergedRegion(region);
	}

	public int addMergedRegionUnsafe(CellRangeAddress region) {
		return delegateSheet.addMergedRegionUnsafe(region);
	}

	public void validateMergedRegions() {
		delegateSheet.validateMergedRegions();
	}

	public void setVerticallyCenter(boolean value) {
		delegateSheet.setVerticallyCenter(value);
	}

	public void setHorizontallyCenter(boolean value) {
		delegateSheet.setHorizontallyCenter(value);
	}

	public boolean getHorizontallyCenter() {
		return delegateSheet.getHorizontallyCenter();
	}

	public boolean getVerticallyCenter() {
		return delegateSheet.getVerticallyCenter();
	}

	public void removeMergedRegion(int index) {
		delegateSheet.removeMergedRegion(index);
	}

	public void removeMergedRegions(Collection<Integer> indices) {
		delegateSheet.removeMergedRegions(indices);
	}

	public int getNumMergedRegions() {
		return delegateSheet.getNumMergedRegions();
	}

	public CellRangeAddress getMergedRegion(int index) {
		return delegateSheet.getMergedRegion(index);
	}

	public List<CellRangeAddress> getMergedRegions() {
		return delegateSheet.getMergedRegions();
	}

	public Iterator<Row> rowIterator() {
		return delegateSheet.rowIterator();
	}

	public void setForceFormulaRecalculation(boolean value) {
		delegateSheet.setForceFormulaRecalculation(value);
	}

	public boolean getForceFormulaRecalculation() {
		return delegateSheet.getForceFormulaRecalculation();
	}

	public void setAutobreaks(boolean value) {
		delegateSheet.setAutobreaks(value);
	}

	public void setDisplayGuts(boolean value) {
		delegateSheet.setDisplayGuts(value);
	}

	public void setDisplayZeros(boolean value) {
		delegateSheet.setDisplayZeros(value);
	}

	public boolean isDisplayZeros() {
		return delegateSheet.isDisplayZeros();
	}

	public void setFitToPage(boolean value) {
		delegateSheet.setFitToPage(value);
	}

	public void setRowSumsBelow(boolean value) {
		delegateSheet.setRowSumsBelow(value);
	}

	public void setRowSumsRight(boolean value) {
		delegateSheet.setRowSumsRight(value);
	}

	public boolean getAutobreaks() {
		return delegateSheet.getAutobreaks();
	}

	public boolean getDisplayGuts() {
		return delegateSheet.getDisplayGuts();
	}

	public boolean getFitToPage() {
		return delegateSheet.getFitToPage();
	}

	public boolean getRowSumsBelow() {
		return delegateSheet.getRowSumsBelow();
	}

	public boolean getRowSumsRight() {
		return delegateSheet.getRowSumsRight();
	}

	public boolean isPrintGridlines() {
		return delegateSheet.isPrintGridlines();
	}

	public void setPrintGridlines(boolean show) {
		delegateSheet.setPrintGridlines(show);
	}

	public boolean isPrintRowAndColumnHeadings() {
		return delegateSheet.isPrintRowAndColumnHeadings();
	}

	public void setPrintRowAndColumnHeadings(boolean show) {
		delegateSheet.setPrintRowAndColumnHeadings(show);
	}

	public PrintSetup getPrintSetup() {
		return delegateSheet.getPrintSetup();
	}

	public Header getHeader() {
		return delegateSheet.getHeader();
	}

	public Footer getFooter() {
		return delegateSheet.getFooter();
	}

	public void setSelected(boolean value) {
		delegateSheet.setSelected(value);
	}

	public double getMargin(short margin) {
		return delegateSheet.getMargin(margin);
	}

	public void setMargin(short margin, double size) {
		delegateSheet.setMargin(margin, size);
	}

	public boolean getProtect() {
		return delegateSheet.getProtect();
	}

	public void protectSheet(String password) {
		delegateSheet.protectSheet(password);
	}

	public boolean getScenarioProtect() {
		return delegateSheet.getScenarioProtect();
	}

	public void setZoom(int scale) {
		delegateSheet.setZoom(scale);
	}

	public short getTopRow() {
		return delegateSheet.getTopRow();
	}

	public short getLeftCol() {
		return delegateSheet.getLeftCol();
	}

	public void showInPane(int toprow, int leftcol) {
		delegateSheet.showInPane(toprow, leftcol);
	}

	public void shiftRows(int startRow, int endRow, int n) {
		delegateSheet.shiftRows(startRow, endRow, n);
	}

	public void shiftRows(int startRow, int endRow, int n, boolean copyRowHeight, boolean resetOriginalRowHeight) {
		delegateSheet.shiftRows(startRow, endRow, n, copyRowHeight, resetOriginalRowHeight);
	}

	public void shiftColumns(int startColumn, int endColumn, int n) {
		delegateSheet.shiftColumns(startColumn, endColumn, n);
	}

	public void createFreezePane(int colSplit, int rowSplit, int leftmostColumn, int topRow) {
		delegateSheet.createFreezePane(colSplit, rowSplit, leftmostColumn, topRow);
	}

	public void createFreezePane(int colSplit, int rowSplit) {
		delegateSheet.createFreezePane(colSplit, rowSplit);
	}

	public void createSplitPane(int xSplitPos, int ySplitPos, int leftmostColumn, int topRow, int activePane) {
		delegateSheet.createSplitPane(xSplitPos, ySplitPos, leftmostColumn, topRow, activePane);
	}

	public PaneInformation getPaneInformation() {
		return delegateSheet.getPaneInformation();
	}

	public void setDisplayGridlines(boolean show) {
		delegateSheet.setDisplayGridlines(show);
	}

	public boolean isDisplayGridlines() {
		return delegateSheet.isDisplayGridlines();
	}

	public void setDisplayFormulas(boolean show) {
		delegateSheet.setDisplayFormulas(show);
	}

	public boolean isDisplayFormulas() {
		return delegateSheet.isDisplayFormulas();
	}

	public void setDisplayRowColHeadings(boolean show) {
		delegateSheet.setDisplayRowColHeadings(show);
	}

	public boolean isDisplayRowColHeadings() {
		return delegateSheet.isDisplayRowColHeadings();
	}

	public void setRowBreak(int row) {
		delegateSheet.setRowBreak(row);
	}

	public boolean isRowBroken(int row) {
		return delegateSheet.isRowBroken(row);
	}

	public void removeRowBreak(int row) {
		delegateSheet.removeRowBreak(row);
	}

	public int[] getRowBreaks() {
		return delegateSheet.getRowBreaks();
	}

	public int[] getColumnBreaks() {
		return delegateSheet.getColumnBreaks();
	}

	public void setColumnBreak(int column) {
		delegateSheet.setColumnBreak(column);
	}

	public boolean isColumnBroken(int column) {
		return delegateSheet.isColumnBroken(column);
	}

	public void removeColumnBreak(int column) {
		delegateSheet.removeColumnBreak(column);
	}

	public void setColumnGroupCollapsed(int columnNumber, boolean collapsed) {
		delegateSheet.setColumnGroupCollapsed(columnNumber, collapsed);
	}

	public void groupColumn(int fromColumn, int toColumn) {
		delegateSheet.groupColumn(fromColumn, toColumn);
	}

	public void ungroupColumn(int fromColumn, int toColumn) {
		delegateSheet.ungroupColumn(fromColumn, toColumn);
	}

	public void groupRow(int fromRow, int toRow) {
		delegateSheet.groupRow(fromRow, toRow);
	}

	public void ungroupRow(int fromRow, int toRow) {
		delegateSheet.ungroupRow(fromRow, toRow);
	}

	public void setRowGroupCollapsed(int row, boolean collapse) {
		delegateSheet.setRowGroupCollapsed(row, collapse);
	}

	public void setDefaultColumnStyle(int column, CellStyle style) {
		delegateSheet.setDefaultColumnStyle(column, style);
	}

	public void autoSizeColumn(int column) {
		delegateSheet.autoSizeColumn(column);
	}

	public void autoSizeColumn(int column, boolean useMergedCells) {
		delegateSheet.autoSizeColumn(column, useMergedCells);
	}

	public Comment getCellComment(CellAddress ref) {
		return delegateSheet.getCellComment(ref);
	}

	public Map<CellAddress, ? extends Comment> getCellComments() {
		return delegateSheet.getCellComments();
	}

	public Drawing<?> getDrawingPatriarch() {
		return delegateSheet.getDrawingPatriarch();
	}

	public Drawing<?> createDrawingPatriarch() {
		return delegateSheet.createDrawingPatriarch();
	}

	public Workbook getWorkbook() {
		return delegateSheet.getWorkbook();
	}

	public String getSheetName() {
		return delegateSheet.getSheetName();
	}

	public boolean isSelected() {
		return delegateSheet.isSelected();
	}

	public CellRange<? extends Cell> setArrayFormula(String formula, CellRangeAddress range) {
		return delegateSheet.setArrayFormula(formula, range);
	}

	public CellRange<? extends Cell> removeArrayFormula(Cell cell) {
		return delegateSheet.removeArrayFormula(cell);
	}

	public DataValidationHelper getDataValidationHelper() {
		return delegateSheet.getDataValidationHelper();
	}

	public List<? extends DataValidation> getDataValidations() {
		return delegateSheet.getDataValidations();
	}

	public void addValidationData(DataValidation dataValidation) {
		delegateSheet.addValidationData(dataValidation);
	}

	public AutoFilter setAutoFilter(CellRangeAddress range) {
		return delegateSheet.setAutoFilter(range);
	}

	public SheetConditionalFormatting getSheetConditionalFormatting() {
		return delegateSheet.getSheetConditionalFormatting();
	}

	public CellRangeAddress getRepeatingRows() {
		return delegateSheet.getRepeatingRows();
	}

	public CellRangeAddress getRepeatingColumns() {
		return delegateSheet.getRepeatingColumns();
	}

	public void setRepeatingRows(CellRangeAddress rowRangeRef) {
		delegateSheet.setRepeatingRows(rowRangeRef);
	}

	public void setRepeatingColumns(CellRangeAddress columnRangeRef) {
		delegateSheet.setRepeatingColumns(columnRangeRef);
	}

	public int getColumnOutlineLevel(int columnIndex) {
		return delegateSheet.getColumnOutlineLevel(columnIndex);
	}

	public Hyperlink getHyperlink(int row, int column) {
		return delegateSheet.getHyperlink(row, column);
	}

	public Hyperlink getHyperlink(CellAddress addr) {
		return delegateSheet.getHyperlink(addr);
	}

	public List<? extends Hyperlink> getHyperlinkList() {
		return delegateSheet.getHyperlinkList();
	}

	public CellAddress getActiveCell() {
		return delegateSheet.getActiveCell();
	}

	public void setActiveCell(CellAddress address) {
		delegateSheet.setActiveCell(address);
	}

	protected SXSSFSheet delegateSheet = null;
	private CellStyle headerStyle = null;
	private CellStyle bodyStyle = null;
	protected AutoFixStyleWorkbook workBook = null;

	public AutoStyleSheet(AutoFixStyleWorkbook workBook, SXSSFSheet delegateSheet, CellStyle headerStyle,
			CellStyle bodyStyle) {
		this.workBook = workBook;
		this.delegateSheet = delegateSheet;
		this.delegateSheet.trackAllColumnsForAutoSizing();
		this.headerStyle = headerStyle;
		this.bodyStyle = bodyStyle;
	}

	public Row createRow(int rownum) {
		Row row = delegateSheet.createRow(rownum);

		Integer rowsize = this.workBook.rowMax.get(this.getSheetName());
		if (rowsize == null || rowsize < rownum) {
			workBook.rowMax.put(this.getSheetName(), rownum);
		}

		if (rownum == 0) {
			return new AutoStyleRow(workBook, this, row, headerStyle);
		} else {
			return new AutoStyleRow(workBook, this, row, bodyStyle);
		}
	}
}
