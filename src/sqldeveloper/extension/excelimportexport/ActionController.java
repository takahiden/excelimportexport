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

import java.util.ArrayList;
import java.util.List;

import javax.swing.JOptionPane;

import oracle.dbtools.raptor.navigator.impl.ObjectNode;
import oracle.dbtools.raptor.utils.DBObject;
import oracle.ide.Context;
import oracle.ide.Ide;
import oracle.ide.controller.Controller;
import oracle.ide.controller.IdeAction;
import oracle.ide.model.Element;
import oracle.ide.view.View;
import oracle.javatools.editor.BasicEditorPane;
import oracle.javatools.editor.BasicEditorPaneContainer;

public class ActionController implements Controller {
	public static int ACTION_EXPORT_CMD = Ide.findOrCreateCmdID("Action_Export_ID");
	public static int ACTION_IMPORT_CMD = Ide.findOrCreateCmdID("Action_Import_ID");

	public static String accessedDir = null;

	@Override
	public boolean handleEvent(IdeAction action, Context context) {
		int cmdId = action.getCommandId();
		if (ACTION_EXPORT_CMD == cmdId) {
			doExportAction(context);
			return true;
		}
		if (ACTION_IMPORT_CMD == cmdId) {
			doImportAction(context);
			return true;
		}
		return false;
	}

	@Override
	public boolean update(IdeAction action, Context context) {
		action.setEnabled(true);
		return true;
	}

	private void doExportAction(Context context) {
		DBObject dbObject = new DBObject(context.getNode());

		String connectionName = (String) context.getProperty("Connections.db_name");
		if (connectionName == null) {
			connectionName = dbObject.getConnectionName();
		}
		// 接続していない場合は、アラートを出して終了
		if (connectionName == null) {
			JOptionPane.showInternalMessageDialog(Ide.getMainWindow().getContentPane(),
					ExtensionResources.format("ERROR_NOT_CONNECT"), ExtensionResources.format("ERROR_TITLE"),
					JOptionPane.WARNING_MESSAGE);
			return;
		}

		StringBuilder strQuery = new StringBuilder("");
		List<String> sequenceNames = new ArrayList<String>();
		Element[] els = context.getSelection();
		for (Element el : els) {
			if (el instanceof ObjectNode) {
				ObjectNode objectEl = (ObjectNode) el;
				if ("TABLE".equals(objectEl.getObjectType())) {
					strQuery.append("SELECT * FROM " + objectEl.getObjectName()
							+ PKListUtil.getPKColumnOrderBy(connectionName, objectEl.getObjectName()) + ";"
							+ System.lineSeparator());
				} else if ("VIEW".equals(objectEl.getObjectType())) {
					strQuery.append("SELECT * FROM " + objectEl.getObjectName() + ";" + System.lineSeparator());
				} else if ("SEQUENCE".equals(objectEl.getObjectType())) {
					sequenceNames.add("'" + objectEl.getObjectName() + "'");
				}
			}
		}

		if (!sequenceNames.isEmpty()) {
			strQuery.append("SELECT * FROM USER_SEQUENCES WHERE SEQUENCE_NAME IN (" + String.join(", ", sequenceNames)
					+ ");" + System.lineSeparator());
		}
		try {
			View view = context.getView();
			BasicEditorPane editor = ((BasicEditorPaneContainer) view).getFocusedEditorPane();
			strQuery.append(editor.getSelectedText() == null ? "" : editor.getSelectedText());
		} catch (Exception e) {
			;
		}

		ExportDialog.createSaveFrame(strQuery.toString(), connectionName);

	}

	private void doImportAction(Context context) {

		String connectionName = (String) context.getProperty("Connections.db_name");
		if (connectionName == null) {
			DBObject dbObject = new DBObject(context.getNode());
			connectionName = dbObject.getConnectionName();
		}
		// 接続していない場合は、アラートを出して終了
		if (connectionName == null) {
			JOptionPane.showInternalMessageDialog(Ide.getMainWindow().getContentPane(),
					ExtensionResources.format("ERROR_NOT_CONNECT"), ExtensionResources.format("ERROR_TITLE"),
					JOptionPane.WARNING_MESSAGE);
			return;
		}
		ImportDialog.createLoadFrame(connectionName);

	}
}
