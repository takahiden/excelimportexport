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
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.util.IOUtils;

import oracle.ide.log.LogManager;

public class LogSheetUtil {

	public static void outputLog(String destFile, List<LogBean> logList) {
		Workbook book = null;

		// read
		FileInputStream in = null;
		try {
			in = new FileInputStream(destFile);
			book = WorkbookFactory.create(in);
		} catch (IOException | InvalidFormatException e) {
			LogMessage("ERROR", e.getMessage());
			return;
		} finally {
			IOUtils.closeQuietly(in);
		}

		Sheet logSheet = book.getSheet("(LOG)");
		if (logSheet == null) {
			logSheet = book.createSheet("(LOG)");
			logSheet.createRow(0).createCell(0).setCellValue("time");
			logSheet.getRow(0).createCell(1).setCellValue("table");
			logSheet.getRow(0).createCell(2).setCellValue("sql");
			logSheet.getRow(0).createCell(3).setCellValue("result");
		}

		for (LogBean log : logList) {
			Row row = logSheet.createRow(logSheet.getLastRowNum() + 1);
			CreationHelper createHelper = book.getCreationHelper();
			CellStyle cellStyle = book.createCellStyle();
			cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("yyyy/mm/dd hh:mm:ss"));
			row.createCell(0).setCellStyle(cellStyle);
			row.getCell(0).setCellValue(log.getTime());
			row.createCell(1).setCellValue(log.getTableName());
			row.createCell(2).setCellValue(log.getSql());
			row.createCell(3).setCellValue(log.getResult());
		}
		logSheet.setColumnWidth(0, 5000);
		logSheet.setColumnWidth(1, 5000);
		logSheet.setColumnWidth(2, 20960);
		// write
		FileOutputStream out = null;
		try {
			out = new FileOutputStream(destFile);
			book.write(out);
		} catch (IOException e) {
			LogMessage("ERROR", e.getMessage());
		} finally {
			IOUtils.closeQuietly(out);
		}
	}

	private static final void LogMessage(String level, String msg) {
		DateFormat dateFormat = new SimpleDateFormat("HH:mm:ss");
		Date date = new Date();
		String currentTime = "[" + dateFormat.format(date) + "] ";
		LogManager.getLogManager().getMsgPage().log("EXPORT " + currentTime + level + ": " + msg + "\n");
	}

}
