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

import java.awt.Dimension;
import java.awt.FlowLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.sql.Connection;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTextArea;
import javax.swing.SwingWorker;
import javax.swing.filechooser.FileNameExtensionFilter;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.dbunit.database.DatabaseConfig;
import org.dbunit.database.DatabaseConnection;
import org.dbunit.database.IDatabaseConnection;
import org.dbunit.dataset.excel.XlsDataSetWriter;

import com.foundationdb.sql.StandardException;
import com.foundationdb.sql.parser.FromTable;
import com.foundationdb.sql.parser.SQLParser;
import com.foundationdb.sql.parser.StatementNode;

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
import oracle.dbtools.raptor.utils.Connections;
import oracle.ide.log.LogManager;
import oracle.javatools.db.DBException;

public class ExportDialog extends JFrame implements ActionListener {

	private static final String btnSave_title = "Save";

	private static final int window_height = 470;
	private static final int window_width = 950;

	private static final long serialVersionUID = 1L;
	JTextArea tf01 = new JTextArea("");

	JCheckBox logOutCheck = null;
	String connectionName = null;

	/**
	 * テスト起動主処理（ウィンドウ表示）
	 */
	public static void main(final String[] args) {

		createSaveFrame("select 1 from dual", "aaa");
	}

	/**
	 * ダイアログ呼び出し
	 * 
	 * @param sql
	 *            SQL文
	 * @param connectionName
	 *            接続先名
	 */
	public static void createSaveFrame(String sql, String connectionName) {
		final ExportDialog f = new ExportDialog();
		f.setTitle(ExtensionResources.format("EXPORT_TITLE"));
		f.tf01.setText("" + sql);
		f.connectionName = connectionName;
		f.setSize(ExportDialog.window_width, ExportDialog.window_height);
		// f.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

		f.setVisible(true);
	}

	JButton btnSave = null;

	/**
	 * コンストラクタ
	 */
	public ExportDialog() {
		btnSave = new JButton(ExportDialog.btnSave_title);
		btnSave.setPreferredSize(new Dimension(200, 30));
		btnSave.addActionListener(this);
		JPanel pane = new JPanel();
		pane.setLayout(new FlowLayout(FlowLayout.CENTER, 1000, 10));
		JScrollPane scrollpane = new JScrollPane(tf01);
		scrollpane.setPreferredSize(new Dimension(910, 300));
		pane.add(scrollpane);
		logOutCheck = new JCheckBox(ExtensionResources.format("LOG_OUTPUT"));
		pane.add(logOutCheck);
		pane.add(btnSave);
		getContentPane().add(pane);
		this.addWindowListener(new WinAdapter());
	}

	/**
	 * ウィンドウを閉じる
	 *
	 */
	class WinAdapter extends WindowAdapter {
		@Override
		public void windowClosing(final WindowEvent we) {
			dispose();
		}
	}

	/**
	 * Saveボタンクリック時処理
	 */
	public void actionPerformed(final ActionEvent ae) {

		if (ae.getActionCommand() == ExportDialog.btnSave_title) {
			String outFilePath = this.writefile();
			if (outFilePath == null) {
				return;
			}
			createExportExcel(this.tf01.getText(), this.connectionName, outFilePath);
		}
	}

	/**
	 * 保存用ファイル選択ダイアログ表示
	 * 
	 * @return 対象ファイル名
	 */
	String writefile() {
		String fullpath = null;

		JFileChooser fc = ActionController.accessedDir == null ? new JFileChooser()
				: new JFileChooser(new File(ActionController.accessedDir));
		fc.setFileSelectionMode(JFileChooser.FILES_ONLY);
		FileNameExtensionFilter ff = new FileNameExtensionFilter(ExtensionResources.format("FILE_FORMAT_NAME"), "xlsx");
		fc.addChoosableFileFilter(ff);
		fc.setFileFilter(ff);
		File file;
		while (true) {
			if (fc.showSaveDialog(this) != JFileChooser.APPROVE_OPTION) {
				return null;
			}
			file = fc.getSelectedFile();
			if (file.toString().substring(file.toString().length() - 5).equalsIgnoreCase(".xlsx")) {
				fullpath = file.getAbsolutePath();
			} else {
				fullpath = file.getAbsolutePath() + ".xlsx";
				file = new File(fullpath);
			}

			if (file.exists()) {
				String msg = ExtensionResources.format("FILE_ALREADY_EXISTS", file.getName());
				switch (JOptionPane.showConfirmDialog(this, msg, ExtensionResources.format("CONFIRM_TITLE"),
						JOptionPane.YES_NO_CANCEL_OPTION)) {
				case JOptionPane.YES_OPTION:
					ActionController.accessedDir = file.getParent();
					return fullpath;
				case JOptionPane.NO_OPTION:
					continue;
				case JOptionPane.CANCEL_OPTION:
					return null;
				}
				break;
			} else {
				break;
			}
		}

		return fullpath;
	}

	protected void createExportExcel(String strQuery, String connectionName, String destFilePath) {
		List<String> queryList = new ArrayList<String>();
		List<String> tableList = new ArrayList<String>();

		if (strQuery != null && strQuery.length() > 0) {
			String[] sa = strQuery.split(";");
			for (String s : sa) {
				s = s.trim();
				if (s.length() > 0) {
					queryList.add(s);
				}
			}
		}

		if (queryList.isEmpty()) {
			JOptionPane.showInternalMessageDialog(getContentPane(), ExtensionResources.format("ERROR_EXPORT_NO_QUERY"),
					ExtensionResources.format("ERROR_TITLE"), JOptionPane.WARNING_MESSAGE);
			return;
		}

		if (!Connections.getInstance().isConnectionOpen(connectionName)) {
			JOptionPane.showInternalMessageDialog(getContentPane(), ExtensionResources.format("ERROR_NOT_CONNECT"),
					ExtensionResources.format("ERROR_TITLE"), JOptionPane.WARNING_MESSAGE);
			return;
		}

		// xlsx形式で出力できるようにする
		XlsDataSetWriter writer = new XlsDataSetWriter() {
			@Override
			public Workbook createWorkbook() {
				return new XSSFWorkbook();
			}
		};

		// ＳＱＬのアクセステーブル名を取得し、シート名につかう
		for (int i = 0; i < queryList.size(); i++) {
			String query = queryList.get(i);
			String tableName = getFromTableName(query);
			if (tableName == null || tableName.length() <= 0 || tableList.contains(tableName)) {
				tableList.add("SQL" + String.format("%03d", ++i));
			} else {
				tableList.add(tableName);
			}
		}
		btnSave.setEnabled(false);
		thread(connectionName, destFilePath, queryList, tableList, writer, this);

	}

	protected void thread(String connectionName, String destFilePath, List<String> queryList, List<String> tableList,
			XlsDataSetWriter writer, ExportDialog parent) {
		new SwingWorker<Object, Object>() {

			@Override
			protected Object doInBackground() throws Exception {

				IDatabaseConnection con;
				QueryDataSetEx partialDataSet = null;
				List<LogBean> logList = new ArrayList<LogBean>();
				try {
					con = new DatabaseConnection(getConnection(connectionName));
					DatabaseConfig config = con.getConfig();
					config.setProperty(DatabaseConfig.PROPERTY_DATATYPE_FACTORY,
							new OracleDataTypeFactoryEx(destFilePath));
					String[] types = new String[] { "TABLE", "VIEW", "MATERIALIZED VIEW" };
					config.setProperty(DatabaseConfig.PROPERTY_TABLE_TYPE, types);
					partialDataSet = new QueryDataSetEx(con, parent, logList);
					for (int i = 0; i < queryList.size(); i++) {
						String query = queryList.get(i);

						partialDataSet.addTable(tableList.get(i), query);
						LogMessage("DEBUG", "query : " + query);
					}

					writer.write(partialDataSet, new FileOutputStream(destFilePath));
				} catch (Exception e) {
					if (QueryDataSetEx.CANCEL_FLG.equals(e.getMessage())) {
						LogMessage("INFO", " export cancel.");
						LogManager.getLogManager().getMsgPage().log("CANCELED " + "\n");
						JOptionPane.showInternalMessageDialog(getContentPane(),
								ExtensionResources.format("CANCEL_EXPORT"), ExtensionResources.format("CANCEL_TITLE"),
								JOptionPane.INFORMATION_MESSAGE);
						return null;
					}
					LogMessage("ERROR", e.getMessage());
					StringWriter errors = new StringWriter();
					e.printStackTrace(new PrintWriter(errors));
					LogMessage("ERROR", errors.toString());
					JOptionPane.showInternalMessageDialog(getContentPane(), ExtensionResources.format("ERROR_EXPORT"),
							ExtensionResources.format("ERROR_TITLE"), JOptionPane.ERROR_MESSAGE);
					return null;
				} finally {
					if (partialDataSet != null) {
						try {
							partialDataSet.close();
						} catch (IOException e) {
							;
						}
					}
				}

				if (logOutCheck.isSelected()) {
					// log
					LogSheetUtil.outputLog(destFilePath, logList);
				}

				FixStyleUtil.fixStyle(destFilePath);

				LogMessage("INFO", " successfully! [" + destFilePath + "]");
				LogManager.getLogManager().getMsgPage().log("CALLED " + "\n");
				JOptionPane.showInternalMessageDialog(getContentPane(), ExtensionResources.format("SUCCESS_EXPORT"),
						ExtensionResources.format("SUCCESS_TITLE"), JOptionPane.INFORMATION_MESSAGE);
				return null;
			}

			@Override
			protected void done() {
				btnSave.setEnabled(true);
			}
		}.execute();
	}

	public Connection getConnection(String connectionName) throws DBException {
		return Connections.getInstance().getConnection(connectionName);
	}

	private static final void LogMessage(String level, String msg) {
		DateFormat dateFormat = new SimpleDateFormat("HH:mm:ss");
		Date date = new Date();
		String currentTime = "[" + dateFormat.format(date) + "] ";
		LogManager.getLogManager().getMsgPage().log("EXPORT " + currentTime + level + ": " + msg + "\n");
	}

	protected String getFromTableName(String query) {
		SQLParser parser = new SQLParser();
		QueryVisitor qVis = new QueryVisitor();

		StatementNode stmt = null;
		try {
			stmt = parser.parseStatement(query);
			stmt.accept(qVis);
		} catch (StandardException e) {
			return null;
		}

		for (FromTable table : qVis.fromList) {
			String tableName;
			try {
				tableName = table.getTableName().toString();
			} catch (StandardException e) {
				return null;
			}
			if (tableName != null) {
				tableName = tableName.toUpperCase();
			}
			return tableName;
		}
		return null;
	}
}