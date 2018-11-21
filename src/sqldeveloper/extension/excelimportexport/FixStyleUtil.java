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

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.POIXMLProperties;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import oracle.ide.log.LogManager;

public class FixStyleUtil {

	public static void main(String[] args) {
		fixStyle("C:/sqldeveloperWork/cccc.xlsx");

	}

	public static void fixStyle(String destFile) {
		XSSFWorkbook book = null;
		FileInputStream in = null;
		FileOutputStream out = null;
		try {
			// read
			in = new FileInputStream(destFile);
			book = new XSSFWorkbook(in);

			POIXMLProperties xmlProps = book.getProperties();
			POIXMLProperties.CoreProperties coreProps = xmlProps.getCoreProperties();

			String username = System.getProperty("user.name");
			if (username == null) {
				username = "author";
			}
			coreProps.setCreator(username);

			CellStyle headerStyle = book.createCellStyle();
			headerStyle.setBorderBottom(CellStyle.BORDER_DOUBLE);
			XSSFFont headerFont = book.createFont();
			headerFont.setBold(true);

			for (int i = 0; i < book.getNumberOfSheets(); i++) {
				XSSFSheet sheet = book.getSheetAt(i);
				if ("(LOG)".equals(sheet.getSheetName())) {
					continue;
				}
				int lastCell = sheet.getRow(0).getLastCellNum();
				for (int col = 0; col < lastCell; col++) {
					if (col == 0) {
						XSSFFont firstRowFont = sheet.getRow(0).getCell(col).getCellStyle().getFont();
//						headerFont.setFontName("Meiryo UI");
						headerFont.setFontHeightInPoints(firstRowFont.getFontHeightInPoints());
						headerStyle.setFont(headerFont);
					}
					sheet.autoSizeColumn(col);
					sheet.getRow(0).getCell(col).setCellStyle(headerStyle);
				}
			}

			// write
			out = new FileOutputStream(destFile);
			book.write(out);

		} catch (IOException e) {
			LogMessage("ERROR", e.getMessage());
			return;
		} finally {
			IOUtils.closeQuietly(in);
			IOUtils.closeQuietly(out);
			IOUtils.closeQuietly(book);
		}
	}

	private static final void LogMessage(String level, String msg) {
		DateFormat dateFormat = new SimpleDateFormat("HH:mm:ss");
		Date date = new Date();
		String currentTime = "[" + dateFormat.format(date) + "] ";
		LogManager.getLogManager().getMsgPage().log("EXPORT " + currentTime + level + ": " + msg + "\n");
	}
}
