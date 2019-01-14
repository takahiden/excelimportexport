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

import java.sql.SQLException;
import java.util.ArrayList;
import java.util.List;

import org.dbunit.database.statement.IPreparedBatchStatement;
import org.dbunit.dataset.datatype.DataType;
import org.dbunit.dataset.datatype.TypeCastException;

public class IPreparedBatchStatementEx implements IPreparedBatchStatement {

	private IPreparedBatchStatement delegate;

	private boolean skip = true;
	private List<Object> valueList = new ArrayList<Object>();
	private List<DataType> dataTypeList = new ArrayList<DataType>();

	public IPreparedBatchStatementEx(IPreparedBatchStatement delegate) {
		this.delegate = delegate;
	}

	@Override
	public void addBatch() throws SQLException {
		if (!skip) {
			for (int i = 0; i < dataTypeList.size(); i++) {
				try {
					delegate.addValue(valueList.get(i), dataTypeList.get(i));
				} catch (TypeCastException e) {
					throw new SQLException(e);
				}
			}
			delegate.addBatch();
		}
		skip = true;
		valueList.clear();
		dataTypeList.clear();
	}

	@Override
	public void addValue(Object value, DataType dataType) throws TypeCastException, SQLException {
		if (!(value == null || (value instanceof String && "".equals(value)))) {
			skip = false;
		}
		valueList.add(value);
		dataTypeList.add(dataType);
	}

	@Override
	public void clearBatch() throws SQLException {
		skip = true;
		valueList.clear();
		dataTypeList.clear();
		delegate.clearBatch();
	}

	@Override
	public void close() throws SQLException {
		skip = true;
		valueList.clear();
		dataTypeList.clear();
		delegate.close();
	}

	@Override
	public int executeBatch() throws SQLException {
		skip = true;
		valueList.clear();
		dataTypeList.clear();
		return delegate.executeBatch();
	}

}
