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

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Spliterator;
import java.util.function.Consumer;

import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.formula.udf.UDFFinder;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.SheetVisibility;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class AutoFixStyleWorkbook implements Workbook {

	private XSSFWorkbook original = null;
	private SXSSFWorkbook delegate = null;
	private File tempFile = null;
	private CellStyle headerStyle = null;
	private CellStyle bodyStyle = null;

	public AutoFixStyleWorkbook(XSSFWorkbook original, SXSSFWorkbook delegate, File tempFile) {
		this.original = original;
		this.delegate = delegate;
		this.tempFile = tempFile;

		headerStyle = delegate.createCellStyle();
		Font headerFont = delegate.createFont();
		headerFont.setFontName("Serif");
		headerFont.setBold(true);
		headerStyle.setFont(headerFont);
		headerStyle.setBorderBottom(BorderStyle.DOUBLE);

		bodyStyle = delegate.createCellStyle();
		Font bodyFont = delegate.createFont();
		bodyFont.setFontName("Serif");
		bodyStyle.setFont(bodyFont);
	}

	public Sheet createSheet() {
		return new AutoStyleSheet(this, delegate.createSheet(), headerStyle, bodyStyle);
	}

	public Sheet createSheet(String sheetname) {
		return new AutoStyleSheet(this, delegate.createSheet(sheetname), headerStyle, bodyStyle);
	}

	public Map<String, Integer> colMax = new HashMap<String, Integer>();
	public Map<String, Integer> rowMax = new HashMap<String, Integer>();

	public void write(OutputStream stream) throws IOException {
		Iterator<Sheet> site = delegate.sheetIterator();
		while (site.hasNext()) {
			Sheet sheet = site.next();
			String sheetName = sheet.getSheetName();
			Integer colsize = colMax.get(sheetName);
			for (int i = 0; i <= colsize; i++) {
				sheet.autoSizeColumn(i, true);
				sheet.setColumnWidth(i, sheet.getColumnWidth(i) + 256 * 3);
			}
			Integer rowsize = rowMax.get(sheetName) + 1;
			String ref = CellReference.convertNumToColString(colsize) + rowsize;
			original.getSheet(sheetName).getCTWorksheet().getDimension().setRef("A1:" + ref);
		}
		ByteArrayOutputStream baos = new ByteArrayOutputStream();
		ZipSecureFile.setMinInflateRatio(0.001);
		delegate.write(baos);
		stream.write(baos.toByteArray());
		baos.close();
	}

	public void close() throws IOException {
		if (delegate != null) {
			try {
				delegate.dispose();
				delegate.close();
			} catch (IOException e) {
				;
			}
		}
		if (original != null) {
			try {
				original.close();
			} catch (IOException e) {
				;
			}
		}
		if (tempFile != null) {
			tempFile.delete();
		}
	}

	public Iterator<Sheet> iterator() {
		return delegate.iterator();
	}

	public void forEach(Consumer<? super Sheet> action) {
		delegate.forEach(action);
	}

	public Spliterator<Sheet> spliterator() {
		return delegate.spliterator();
	}

	public int getActiveSheetIndex() {
		return delegate.getActiveSheetIndex();
	}

	public void setActiveSheet(int sheetIndex) {
		delegate.setActiveSheet(sheetIndex);
	}

	public int getFirstVisibleTab() {
		return delegate.getFirstVisibleTab();
	}

	public void setFirstVisibleTab(int sheetIndex) {
		delegate.setFirstVisibleTab(sheetIndex);
	}

	public void setSheetOrder(String sheetname, int pos) {
		delegate.setSheetOrder(sheetname, pos);
	}

	public void setSelectedTab(int index) {
		delegate.setSelectedTab(index);
	}

	public void setSheetName(int sheet, String name) {
		delegate.setSheetName(sheet, name);
	}

	public String getSheetName(int sheet) {
		return delegate.getSheetName(sheet);
	}

	public int getSheetIndex(String name) {
		return delegate.getSheetIndex(name);
	}

	public int getSheetIndex(Sheet sheet) {
		return delegate.getSheetIndex(sheet);
	}

	public Sheet cloneSheet(int sheetNum) {
		return delegate.cloneSheet(sheetNum);
	}

	public Iterator<Sheet> sheetIterator() {
		return delegate.sheetIterator();
	}

	public int getNumberOfSheets() {
		return delegate.getNumberOfSheets();
	}

	public Sheet getSheetAt(int index) {
		return delegate.getSheetAt(index);
	}

	public Sheet getSheet(String name) {
		return delegate.getSheet(name);
	}

	public void removeSheetAt(int index) {
		delegate.removeSheetAt(index);
	}

	public Font createFont() {
		return delegate.createFont();
	}

	public Font findFont(boolean bold, short color, short fontHeight, String name, boolean italic, boolean strikeout,
			short typeOffset, byte underline) {
		return delegate.findFont(bold, color, fontHeight, name, italic, strikeout, typeOffset, underline);
	}

	public short getNumberOfFonts() {
		return delegate.getNumberOfFonts();
	}

	public int getNumberOfFontsAsInt() {
		return delegate.getNumberOfFontsAsInt();
	}

	public Font getFontAt(short idx) {
		return delegate.getFontAt(idx);
	}

	public Font getFontAt(int idx) {
		return delegate.getFontAt(idx);
	}

	public CellStyle createCellStyle() {
		return delegate.createCellStyle();
	}

	public int getNumCellStyles() {
		return delegate.getNumCellStyles();
	}

	public CellStyle getCellStyleAt(int idx) {
		return delegate.getCellStyleAt(idx);
	}

	public int getNumberOfNames() {
		return delegate.getNumberOfNames();
	}

	public Name getName(String name) {
		return delegate.getName(name);
	}

	public List<? extends Name> getNames(String name) {
		return delegate.getNames(name);
	}

	public List<? extends Name> getAllNames() {
		return delegate.getAllNames();
	}

	public Name getNameAt(int nameIndex) {
		return delegate.getNameAt(nameIndex);
	}

	public Name createName() {
		return delegate.createName();
	}

	public int getNameIndex(String name) {
		return delegate.getNameIndex(name);
	}

	public void removeName(int index) {
		delegate.removeName(index);
	}

	public void removeName(String name) {
		delegate.removeName(name);
	}

	public void removeName(Name name) {
		delegate.removeName(name);
	}

	public int linkExternalWorkbook(String name, Workbook workbook) {
		return delegate.linkExternalWorkbook(name, workbook);
	}

	public void setPrintArea(int sheetIndex, String reference) {
		delegate.setPrintArea(sheetIndex, reference);
	}

	public void setPrintArea(int sheetIndex, int startColumn, int endColumn, int startRow, int endRow) {
		delegate.setPrintArea(sheetIndex, startColumn, endColumn, startRow, endRow);
	}

	public String getPrintArea(int sheetIndex) {
		return delegate.getPrintArea(sheetIndex);
	}

	public void removePrintArea(int sheetIndex) {
		delegate.removePrintArea(sheetIndex);
	}

	public MissingCellPolicy getMissingCellPolicy() {
		return delegate.getMissingCellPolicy();
	}

	public void setMissingCellPolicy(MissingCellPolicy missingCellPolicy) {
		delegate.setMissingCellPolicy(missingCellPolicy);
	}

	public DataFormat createDataFormat() {
		return delegate.createDataFormat();
	}

	public int addPicture(byte[] pictureData, int format) {
		return delegate.addPicture(pictureData, format);
	}

	public List<? extends PictureData> getAllPictures() {
		return delegate.getAllPictures();
	}

	public CreationHelper getCreationHelper() {
		return delegate.getCreationHelper();
	}

	public boolean isHidden() {
		return delegate.isHidden();
	}

	public void setHidden(boolean hiddenFlag) {
		delegate.setHidden(hiddenFlag);
	}

	public boolean isSheetHidden(int sheetIx) {
		return delegate.isSheetHidden(sheetIx);
	}

	public boolean isSheetVeryHidden(int sheetIx) {
		return delegate.isSheetVeryHidden(sheetIx);
	}

	public void setSheetHidden(int sheetIx, boolean hidden) {
		delegate.setSheetHidden(sheetIx, hidden);
	}

	public SheetVisibility getSheetVisibility(int sheetIx) {
		return delegate.getSheetVisibility(sheetIx);
	}

	public void setSheetVisibility(int sheetIx, SheetVisibility visibility) {
		delegate.setSheetVisibility(sheetIx, visibility);
	}

	public void addToolPack(UDFFinder toopack) {
		delegate.addToolPack(toopack);
	}

	public void setForceFormulaRecalculation(boolean value) {
		delegate.setForceFormulaRecalculation(value);
	}

	public boolean getForceFormulaRecalculation() {
		return delegate.getForceFormulaRecalculation();
	}

	public SpreadsheetVersion getSpreadsheetVersion() {
		return delegate.getSpreadsheetVersion();
	}

	public int addOlePackage(byte[] oleData, String label, String fileName, String command) throws IOException {
		return delegate.addOlePackage(oleData, label, fileName, command);
	}

}
