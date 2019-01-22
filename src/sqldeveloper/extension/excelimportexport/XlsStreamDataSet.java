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
import java.io.OutputStream;

import org.apache.poi.ss.usermodel.Workbook;
import org.dbunit.dataset.AbstractDataSet;
import org.dbunit.dataset.DataSetException;
import org.dbunit.dataset.DefaultTableIterator;
import org.dbunit.dataset.IDataSet;
import org.dbunit.dataset.ITable;
import org.dbunit.dataset.ITableIterator;
import org.dbunit.dataset.OrderedTableNameMap;
import org.dbunit.dataset.excel.XlsDataSetWriter;

import com.monitorjbl.xlsx.StreamingReader;

public class XlsStreamDataSet extends AbstractDataSet {
	protected Workbook workbook;

	public Workbook getWorkbook() {
		return workbook;
	}

	public static final String DOCUMENT_BUILDER_FACTORY_KEY = "javax.xml.parsers.DocumentBuilderFactory";
	public static final String TRANSFORMER_FACTORY_KEY = "javax.xml.transform.TransformerFactory";
	public static final String XML_INPUT_FACTORY_KEY = "javax.xml.stream.XMLInputFactory";
	public static final String XML_OUTPUT_FACTORY_KEY = "javax.xml.stream.XMLOutputFactory";
	public static final String XML_EVENT_FACTORY_KEY = "javax.xml.stream.XMLEventFactory";

	public XlsStreamDataSet(File file) throws IOException, DataSetException {
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

			workbook = StreamingReader.builder().rowCacheSize(100).bufferSize(4096).open(file);
		} finally {

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

		int sheetCount = workbook.getNumberOfSheets();
		for (int i = 0; i < sheetCount; i++) {
			ITable table = new XlsStreamTable(workbook.getSheetName(i), workbook.getSheetAt(i));
			this._tables.add(table.getTableMetaData().getTableName(), table);
		}
	}

	private final OrderedTableNameMap _tables = super.createTableNameMap();

	public static void write(IDataSet dataSet, OutputStream out) throws IOException, DataSetException {
		new XlsDataSetWriter().write(dataSet, out);
	}

	@SuppressWarnings("unchecked")
	protected ITableIterator createIterator(boolean reversed) throws DataSetException {
		ITable[] tables = (ITable[]) this._tables.orderedValues().toArray(new ITable[0]);
		return new DefaultTableIterator(tables, reversed);
	}
}
