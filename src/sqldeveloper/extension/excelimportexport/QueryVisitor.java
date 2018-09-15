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

import com.foundationdb.sql.StandardException;
import com.foundationdb.sql.parser.CursorNode;
import com.foundationdb.sql.parser.FromList;
import com.foundationdb.sql.parser.GroupByList;
import com.foundationdb.sql.parser.NodeTypes;
import com.foundationdb.sql.parser.OrderByList;
import com.foundationdb.sql.parser.QueryTreeNode;
import com.foundationdb.sql.parser.ResultColumnList;
import com.foundationdb.sql.parser.SelectNode;
import com.foundationdb.sql.parser.ValueNode;
import com.foundationdb.sql.parser.Visitable;
import com.foundationdb.sql.parser.Visitor;

public class QueryVisitor implements Visitor {

	public ResultColumnList resultList;
	public FromList fromList;
	public ValueNode whereClauses;
	public ValueNode havingClauses;
	public GroupByList groupList;
	public OrderByList orderbyList;

	@Override
	public Visitable visit(Visitable visitable) {
		QueryTreeNode node = (QueryTreeNode) visitable;

		switch (node.getNodeType()) {
		case NodeTypes.SELECT_NODE:
			SelectNode sn = (SelectNode) node;
			resultList = sn.getResultColumns();
			fromList = sn.getFromList();
			whereClauses = sn.getWhereClause();
			havingClauses = sn.getHavingClause();
			groupList = sn.getGroupByList();
			break;
		case NodeTypes.CURSOR_NODE:
			orderbyList = ((CursorNode) node).getOrderByList();
			break;
		default:
			break;
		}

		return visitable;
	}

	@Override
	public boolean visitChildrenFirst(Visitable node) {
		return false;
	}

	@Override
	public boolean stopTraversal() {
		return false;
	}

	@Override
	public boolean skipChildren(Visitable node) throws StandardException {
		return false;
	}
}