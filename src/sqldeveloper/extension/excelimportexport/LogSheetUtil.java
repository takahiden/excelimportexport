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

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.io.Writer;
import java.net.URI;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.xml.namespace.QName;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.events.Attribute;
import javax.xml.stream.events.StartElement;
import javax.xml.stream.events.XMLEvent;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackagePartName;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.openxml4j.opc.PackageRelationshipCollection;
import org.apache.poi.openxml4j.opc.PackagingURIHelper;
import org.apache.poi.openxml4j.opc.TargetMode;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.model.ThemesTable;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;

import oracle.ide.log.LogManager;

public class LogSheetUtil {

	public static final String DOCUMENT_BUILDER_FACTORY_KEY = "javax.xml.parsers.DocumentBuilderFactory";
	public static final String TRANSFORMER_FACTORY_KEY = "javax.xml.transform.TransformerFactory";
	public static final String XML_INPUT_FACTORY_KEY = "javax.xml.stream.XMLInputFactory";
	public static final String XML_OUTPUT_FACTORY_KEY = "javax.xml.stream.XMLOutputFactory";
	public static final String XML_EVENT_FACTORY_KEY = "javax.xml.stream.XMLEventFactory";

	public static void outputLog(String destFilePath, List<LogBean> logList) {
		File destFile = new File(destFilePath);

		OPCPackage pkg = null;

		String documentBuilderFactory = System.getProperty(DOCUMENT_BUILDER_FACTORY_KEY);
		String transformerFactory = System.getProperty(TRANSFORMER_FACTORY_KEY);
		String xmlInputFactory = System.getProperty(XML_INPUT_FACTORY_KEY);
		String xmlOutputFactory = System.getProperty(XML_OUTPUT_FACTORY_KEY);
		String xmlEventFactory = System.getProperty(XML_EVENT_FACTORY_KEY);

		try {
			System.setProperty(DOCUMENT_BUILDER_FACTORY_KEY, "org.apache.xerces.jaxp.DocumentBuilderFactoryImpl");
			System.setProperty(TRANSFORMER_FACTORY_KEY,
					"com.sun.org.apache.xalan.internal.xsltc.trax.TransformerFactoryImpl");
			System.setProperty(XML_INPUT_FACTORY_KEY, "com.sun.xml.internal.stream.XMLInputFactoryImpl");
			System.setProperty(XML_OUTPUT_FACTORY_KEY, "com.sun.xml.internal.stream.XMLOutputFactoryImpl");
			System.setProperty(XML_EVENT_FACTORY_KEY, "com.sun.xml.internal.stream.events.XMLEventFactoryImpl");

			pkg = OPCPackage.open(destFile, PackageAccess.READ_WRITE);

			PackagePart logSheetWsPart = getLogSheet(pkg);

			try (StringWriter writer = new StringWriter(); OutputStream os = logSheetWsPart.getOutputStream();) {

				SpreadsheetWriter sw = new SpreadsheetWriter(writer, logList);
				sw.beginSheet();

				sw.insertRow(0);
				sw.createCell(0, "time");
				sw.createCell(1, "table");
				sw.createCell(2, "sql");
				sw.createCell(3, "result");
				sw.endRow();

				for (int rownum = 0; rownum < logList.size(); rownum++) {
					LogBean logBean = logList.get(rownum);
					sw.insertRow(rownum + 1);
					sw.createCell(0, logBean.getTime());
					sw.createCell(1, logBean.getTableName());
					sw.createCell(2, logBean.getSql());
					sw.createCell(3, logBean.getResult());
					sw.createCell(4, logBean.getElse1());
					sw.createCell(5, logBean.getElse2());
					sw.createCell(6, logBean.getElse3());
					sw.createCell(7, logBean.getElse4());
					sw.createCell(8, logBean.getElse5());
					sw.endRow();
				}
				sw.endSheet();
				byte[] arr = writer.toString().getBytes("UTF-8");

				os.write(arr);
			}

		} catch (Throwable ex) {
			StringWriter sw = new StringWriter();
			PrintWriter pw = new PrintWriter(sw);
			ex.printStackTrace(pw);
			LogMessage("ERROR", sw.toString());
		} finally {
			try {
				pkg.close();
			} catch (Exception e) {
				StringWriter sw = new StringWriter();
				PrintWriter pw = new PrintWriter(sw);
				e.printStackTrace(pw);
				LogMessage("ERROR", e.getMessage());
			}

			if (documentBuilderFactory != null) {
				System.setProperty(DOCUMENT_BUILDER_FACTORY_KEY, documentBuilderFactory);
			}
			if (transformerFactory != null) {
				System.setProperty(TRANSFORMER_FACTORY_KEY, transformerFactory);
			}
			if (xmlInputFactory != null) {
				System.setProperty(XML_INPUT_FACTORY_KEY, xmlInputFactory);
			}
			if (xmlOutputFactory != null) {
				System.setProperty(XML_OUTPUT_FACTORY_KEY, xmlOutputFactory);
			}
			if (xmlEventFactory != null) {
				System.setProperty(XML_EVENT_FACTORY_KEY, xmlEventFactory);
			}
		}
	}

	public static class SpreadsheetWriter {
		private final Writer _out;
		private int _rownum;
		private List<LogBean> _logList;

		public SpreadsheetWriter(Writer out, List<LogBean> logList) {
			_out = out;
			_logList = logList;
		}

		public void beginSheet() throws IOException {
			_out.write("<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
					+ "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
			_out.write("<dimension ref=\"A1:I" + (_logList.size() + 1) + "\" />");
			_out.write("<cols>\n");
			_out.write("<col min=\"1\" max=\"1\" width=\"19.75\" bestFit=\"1\" customWidth=\"1\" />\n");
			_out.write("<col min=\"2\" max=\"2\" width=\"12.75\" customWidth=\"1\" />\n");
			_out.write("<col min=\"3\" max=\"3\" width=\"47.5\" customWidth=\"1\" />\n");
			_out.write("<col min=\"4\" max=\"4\" width=\"11.125\" customWidth=\"1\" />\n");
			_out.write("</cols>\n");
			_out.write("<sheetData>\n");
		}

		public void endSheet() throws IOException {
			_out.write("</sheetData>");
			_out.write("</worksheet>");
		}

		/**
		 * Insert a new row
		 *
		 * @param rownum
		 *            0-based row number
		 */
		public void insertRow(int rownum) throws IOException {
			_out.write("<row r=\"" + (rownum + 1) + "\">\n");
			this._rownum = rownum;
		}

		/**
		 * Insert row end marker
		 */
		public void endRow() throws IOException {
			_out.write("</row>\n");
		}

		public void createCell(int columnIndex, String value, int styleIndex) throws IOException {
			String ref = new CellReference(_rownum, columnIndex).formatAsString();
			_out.write("<c r=\"" + ref + "\" t=\"inlineStr\"");
			if (styleIndex != -1)
				_out.write(" s=\"" + styleIndex + "\"");
			_out.write(">");
			_out.write("<is><t>" + (value == null ? "" : value) + "</t></is>");
			_out.write("</c>");
		}

		public void createCell(int columnIndex, String value) throws IOException {
			createCell(columnIndex, value, -1);
		}

		public void createCell(int columnIndex, double value, int styleIndex) throws IOException {
			String ref = new CellReference(_rownum, columnIndex).formatAsString();
			_out.write("<c r=\"" + ref + "\" t=\"n\"");
			if (styleIndex != -1)
				_out.write(" s=\"" + styleIndex + "\"");
			_out.write(">");
			_out.write("<v>" + value + "</v>");
			_out.write("</c>");
		}

		public void createCell(int columnIndex, double value) throws IOException {
			createCell(columnIndex, value, -1);
		}

		public void createCell(int columnIndex, Calendar value, int styleIndex) throws IOException {
			createCell(columnIndex, DateUtil.getExcelDate(value, false), styleIndex);
		}
	}

	public static PackagePart getLogSheet(OPCPackage pkg) throws Exception {

		PackagePart retWsPart = null;

		PackagePart workbookpart = pkg.getPartsByName(Pattern.compile("/xl/workbook.xml")).get(0);
		List<String> sheetNames = new ArrayList<>();
		int logSheetId = 0;
		int maxSheetId = 0;
		int maxSheetNameId = 0;
		URI logUri = null;

		// search log sheet
		try (InputStream is = workbookpart.getInputStream()) {

			String worksheetRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
			PackageRelationshipCollection rs = workbookpart.getRelationshipsByType(worksheetRelationshipType);
			String patternString = ".*sheet(\\d)\\.xml";

			Pattern p = Pattern.compile(patternString);

			for (PackageRelationship r : rs) {
				String target = r.getTargetURI().toString();
				Matcher m = p.matcher(target);
				if (m.find()) {
					String sheetNameId = m.group(1);
					int id = Integer.parseInt(sheetNameId);
					if (maxSheetNameId < id) {
						maxSheetNameId = id;
					}
				}
			}

			XMLEventReader reader = XMLInputFactory.newInstance().createXMLEventReader(is);
			while (reader.hasNext()) {
				XMLEvent event = (XMLEvent) reader.next();

				if (event.isStartElement()) {
					StartElement startElement = (StartElement) event;
					QName startElementName = startElement.getName();
					if (startElementName.getLocalPart().equalsIgnoreCase("sheet")) {
						Attribute attributeName = startElement.getAttributeByName(new QName("name"));
						String sheetName = attributeName.getValue();

						Attribute attributeSheetId = startElement.getAttributeByName(new QName("sheetId"));
						String sheetIdStr = attributeSheetId.getValue();
						int id = Integer.parseInt(sheetIdStr);
						if (maxSheetId < id) {
							maxSheetId = id;
						}
						if ("(LOG)".equals(sheetName)) {
							Attribute attributeRId = startElement.getAttributeByName(new QName(
									"http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id", "r"));
							String rIdStr = attributeRId.getValue();
							PackageRelationship relationShip = workbookpart.getRelationship(rIdStr);
							logUri = relationShip.getTargetURI();
							PackagePartName workSheetUri = PackagingURIHelper.createPartName(logUri);
							retWsPart = pkg.getPart(workSheetUri);

							return retWsPart;
						}
						sheetNames.add(sheetName);
					}
				}
			}
		}

		// if not found log sheet

		logSheetId = (maxSheetId + 1);
		String nsWorksheet = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
		String worksheetRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
		PackagePartName workSheetUri = PackagingURIHelper
				.createPartName("/xl/worksheets/sheet" + (maxSheetNameId + 1) + ".xml");
		PackagePart wsPart = pkg.createPart(workSheetUri, nsWorksheet);

		PackageRelationship relationShip = workbookpart.addRelationship(workSheetUri, TargetMode.INTERNAL,
				worksheetRelationshipType);

		byte[] docBytes = null;
		try (InputStream is = workbookpart.getInputStream(); StringWriter writerDoc = new StringWriter();) {

			DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
			DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
			Document doc = dBuilder.parse(is);

			Node sheets = doc.getElementsByTagName("sheets").item(0);

			Element sheet = doc.createElement("sheet");
			sheet.setAttribute("name", "(LOG)");
			sheet.setAttribute("sheetId", String.valueOf(logSheetId));
			sheet.setAttribute("r:id", relationShip.getId());
			sheets.appendChild(sheet);

			DOMSource source = new DOMSource(doc);
			StreamResult result = new StreamResult(writerDoc);
			TransformerFactory transformerFactory = TransformerFactory.newInstance();
			Transformer transformer = transformerFactory.newTransformer();
			transformer.transform(source, result);
			docBytes = writerDoc.toString().getBytes("UTF-8");
		}

		try (OutputStream osPart = workbookpart.getOutputStream()) {
			osPart.write(docBytes);
		}

		return wsPart;
	}

	public static StylesTable getStylesTable(OPCPackage pkg) throws IOException, InvalidFormatException {
		ArrayList<PackagePart> parts = pkg.getPartsByContentType(XSSFRelation.STYLES.getContentType());
		if (parts.size() == 0)
			return null;

		StylesTable styles = new StylesTable(parts.get(0));
		parts = pkg.getPartsByContentType(XSSFRelation.THEME.getContentType());
		if (parts.size() != 0) {
			styles.setTheme(new ThemesTable(parts.get(0)));
		}
		return styles;
	}

	private static final void LogMessage(String level, String msg) {
		DateFormat dateFormat = new SimpleDateFormat("HH:mm:ss");
		Date date = new Date();
		String currentTime = "[" + dateFormat.format(date) + "] ";
		LogManager.getLogManager().getMsgPage().log("EXPORT " + currentTime + level + ": " + msg + "\n");
	}

}
