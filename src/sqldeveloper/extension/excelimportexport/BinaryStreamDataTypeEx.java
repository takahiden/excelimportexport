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
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.concurrent.atomic.AtomicInteger;

import org.dbunit.dataset.datatype.BinaryStreamDataType;
import org.dbunit.dataset.datatype.TypeCastException;

public class BinaryStreamDataTypeEx extends BinaryStreamDataType {

	public String resultBLobFilePrefix;
	public AtomicInteger counter;
	public File blobDir;

	public BinaryStreamDataTypeEx(String name, int sqlType, File blobDir, String resultBLobFilePrefix,
			AtomicInteger counter) {
		super(name, sqlType);
		this.resultBLobFilePrefix = resultBLobFilePrefix;
		this.counter = counter;
		this.blobDir = blobDir;
	}

	@Override
	public Object getSqlValue(int name, ResultSet resultSet) throws SQLException, TypeCastException {
		Object value = super.getSqlValue(name, resultSet);
		byte[] valueBytes = (byte[]) value;
		if (valueBytes == null) {
			return null;
		}
		if (!blobDir.exists()) {
			blobDir.mkdirs();
		}
		String destFilePath = resultBLobFilePrefix + String.format("%06d", counter.incrementAndGet()) + ".dat";
		try (FileOutputStream fileOutStr = new FileOutputStream(destFilePath)) {
			fileOutStr.write(valueBytes);
		} catch (IOException e) {
			throw new TypeCastException(e);
		}
		return destFilePath;
	}

	@Override
	public void setSqlValue(Object value, int column, PreparedStatement statement)
			throws SQLException, TypeCastException {
		super.setSqlValue(value, column, statement);
	}

}
