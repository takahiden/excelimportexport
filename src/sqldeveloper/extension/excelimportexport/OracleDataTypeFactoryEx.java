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

	public OracleDataTypeFactoryEx(String resultFile) {
		this.resultFile = resultFile;
		this.counter = new AtomicInteger(0);

		String abstPath = new File(this.resultFile).getAbsolutePath();
		int point = abstPath.lastIndexOf(".");
		if (point != -1) {
			resultBLobFilePrefix = abstPath.substring(0, point) + "_BLOB_";
		} else {
			resultBLobFilePrefix = abstPath + "_BLOB_";
		}

	}

	@Override
	public DataType createDataType(int sqlType, String sqlTypeName) throws DataTypeException {
		if ("BLOB".equals(sqlTypeName)) {
			return new BinaryStreamDataTypeEx(sqlTypeName, sqlType, resultBLobFilePrefix, counter);
		} else if (sqlType == Types.DATE) {
			return DataType.VARCHAR;
		} else if (sqlType == Types.TIMESTAMP) {
			return DataType.VARCHAR;
		} else {
			return super.createDataType(sqlType, sqlTypeName);
		}
	}

}
