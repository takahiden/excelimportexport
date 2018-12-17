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
import java.sql.Types;
import java.util.concurrent.atomic.AtomicInteger;

import org.dbunit.dataset.datatype.DataType;
import org.dbunit.dataset.datatype.DataTypeException;
import org.dbunit.ext.oracle.Oracle10DataTypeFactory;

public class OracleDataTypeFactoryEx extends Oracle10DataTypeFactory {
	public AtomicInteger counter;

	public String resultFile;

	public String resultBLobFilePrefix;
	
	public File blobDir;

	public OracleDataTypeFactoryEx(String resultFile) {
		this.resultFile = resultFile;
		this.counter = new AtomicInteger(0);

		File file = new File(this.resultFile);

		File dir = file.getParentFile();
		String fileName = getNameWithoutExtension(file);
		blobDir = new File(dir, fileName + "_BLOB");

		resultBLobFilePrefix = new File(blobDir, fileName + "_BLOB_").getAbsolutePath();
	}

	public String getNameWithoutExtension(File file) {
		String fileName = file.getName();
		int index = fileName.lastIndexOf('.');
		if (index != -1) {
			return fileName.substring(0, index);
		}
		return fileName;
	}

	@Override
	public DataType createDataType(int sqlType, String sqlTypeName) throws DataTypeException {
		if ("BLOB".equals(sqlTypeName)) {
			return new BinaryStreamDataTypeEx(sqlTypeName, sqlType, blobDir, resultBLobFilePrefix, counter);
		} else if (sqlType == Types.DATE) {
			return new StringDateType();
		} else if (sqlType == Types.TIMESTAMP) {
			return new StringTimestampType();
		} else if ("NUMBER".equals(sqlTypeName)) {
			return DataType.VARCHAR;
		} else {
			return super.createDataType(sqlType, sqlTypeName);
		}
	}

}
