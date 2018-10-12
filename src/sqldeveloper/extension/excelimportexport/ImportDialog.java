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
import java.awt.GridLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.ItemEvent;
import java.awt.event.ItemListener;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.io.File;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.sql.Connection;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

import javax.swing.BoxLayout;
import javax.swing.ButtonGroup;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JRadioButton;
import javax.swing.filechooser.FileNameExtensionFilter;

import org.dbunit.database.DatabaseConfig;
import org.dbunit.database.DatabaseConnection;
import org.dbunit.database.IDatabaseConnection;
import org.dbunit.dataset.IDataSet;
import org.dbunit.dataset.ReplacementDataSet;
import org.dbunit.dataset.excel.XlsDataSet;
import org.dbunit.operation.DatabaseOperation;

import oracle.dbtools.raptor.utils.Connections;
import oracle.ide.log.LogManager;
import oracle.javatools.db.DBException;

public class ImportDialog extends JFrame implements ActionListener {

	private static final long serialVersionUID = 1L;

	private static final String btnLoad_title = "Load";

	private static final int window_height = 300;
	private static final int window_width = 600;

	ButtonGroup chkGroup;

	String connectionName = null;
	String schemaName = null;
	boolean deleteBeforeImport = false;

	/**
	 * テスト起動主処理（ウィンドウ表示）
	 */
	public static void main(final String[] args) {
		createLoadFrame("test");
	}

	/**
	 * ダイアログ呼び出し
	 * 
	 * @param sql
	 *            SQL文
	 * @param connectionName
	 *            接続先名
	 */
	public static void createLoadFrame(String connectionName) {
		final ImportDialog f = new ImportDialog();
		f.setTitle(ExtensionResources.format("IMPORT_TITLE"));
		f.connectionName = connectionName;
		f.setSize(ImportDialog.window_width, ImportDialog.window_height);
		// f.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		f.setVisible(true);
	}

	/**
	 * コンストラクタ（部品をセット）
	 */
	public ImportDialog() {

		JButton btnLoad = new JButton(ImportDialog.btnLoad_title);
		btnLoad.setMaximumSize(new Dimension(200, 30));
		btnLoad.addActionListener(this);
		JPanel pane1 = new JPanel();
		pane1.setLayout(new FlowLayout());
		// チェックボックスグループを作成します。
		this.chkGroup = new ButtonGroup();
		// チェックボックスを作成します。
		JRadioButton chk1 = new JRadioButton(ExtensionResources.format("NOT_DELETE_BEFORE_IMPORT"), true);
		JRadioButton chk2 = new JRadioButton(ExtensionResources.format("DELETE_BEFORE_IMPORT"), false);
		chk1.addItemListener(this.new ChkItemAdapter());
		chk2.addItemListener(this.new ChkItemAdapter());
		chkGroup.add(chk1);
		chkGroup.add(chk2);
		pane1.add(chk1);
		pane1.add(chk2);
		JPanel pane2 = new JPanel();
		pane2.setLayout(new BoxLayout(pane2, BoxLayout.Y_AXIS));

		pane2.add(btnLoad);
		GridLayout grid = new GridLayout(2, 1);
		getContentPane().setLayout(grid);
		getContentPane().add(pane1);
		getContentPane().add(pane2);
		btnLoad.setAlignmentX(0.5f);
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
	 * Loadボタンクリック時処理
	 */
	public void actionPerformed(final ActionEvent ae) {

		if (ae.getActionCommand() == ImportDialog.btnLoad_title) {
			String loadFilePath = this.loadfile();
			if (loadFilePath == null) {
				return;
			}
			loadImportExcel(this.deleteBeforeImport, this.connectionName, loadFilePath);
		}
	}

	/**
	 * 読み込み用ファイル選択ダイアログ表示
	 * 
	 * @return 対象ファイル名
	 */
	String loadfile() {
		String fullpath = null;

		JFileChooser fc = ActionController.accessedDir == null ? new JFileChooser()
				: new JFileChooser(new File(ActionController.accessedDir));
		fc.setFileSelectionMode(JFileChooser.FILES_ONLY);
		FileNameExtensionFilter ff = new FileNameExtensionFilter(ExtensionResources.format("FILE_FORMAT_NAME"), "xlsx");
		fc.addChoosableFileFilter(ff);
		fc.setFileFilter(ff);
		File file;

		if (fc.showOpenDialog(this) != JFileChooser.APPROVE_OPTION) {
			return null;
		}
		file = fc.getSelectedFile();

		if (!file.exists()) {
			String msg = ExtensionResources.format("FILE_ALREADY_EXISTS", file.getName());
			JOptionPane.showInternalMessageDialog(getContentPane(), msg, ExtensionResources.format("CONFIRM_TITLE"),
					JOptionPane.WARNING_MESSAGE);
			return null;
		}

		fullpath = file.getAbsolutePath();
		ActionController.accessedDir = file.getParent();

		return fullpath;
	}

	protected void loadImportExcel(boolean deleteBeforeImport, String connectionName, String inputFilePath) {

		if (!Connections.getInstance().isConnectionOpen(connectionName)) {
			JOptionPane.showInternalMessageDialog(getContentPane(), ExtensionResources.format("ERROR_NOT_CONNECT"),
					ExtensionResources.format("ERROR_TITLE"), JOptionPane.WARNING_MESSAGE);
			dispose();
			return;
		}
		IDatabaseConnection con = null;

		try {
			IDataSet dataset = new XlsDataSet(new File(inputFilePath));
			Connection conn = getConnection(connectionName);
			con = new DatabaseConnection(conn, conn.getSchema());
			DatabaseConfig config = con.getConfig();
			config.setProperty(DatabaseConfig.PROPERTY_DATATYPE_FACTORY, new OracleDataTypeFactoryEx(inputFilePath));
			DateFormat df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss.SSS");
			Date now = new Date();
			ReplacementDataSet dataSet = new ReplacementDataSet(dataset);
			dataSet.addReplacementSubstring("SYSDATE", df.format(now));
			dataSet.addReplacementSubstring("SYSTIMESTAMP", df.format(now));
			dataSet.addReplacementSubstring("null", "");
			if (deleteBeforeImport) {
				DatabaseOperation.CLEAN_INSERT.execute(con, dataset);
			} else {
				DatabaseOperation.INSERT.execute(con, dataset);
			}
			con.getConnection().commit();
		} catch (Exception e) {
			try {
				con.getConnection().rollback();
			} catch (Exception e1) {
				;
			}
			LogMessage("ERROR", e.getMessage());
			StringWriter errors = new StringWriter();
			e.printStackTrace(new PrintWriter(errors));
			LogMessage("ERROR", errors.toString());
			JOptionPane.showInternalMessageDialog(getContentPane(), ExtensionResources.format("ERROR_IMPORT"),
					ExtensionResources.format("ERROR_TITLE"), JOptionPane.ERROR_MESSAGE);
			return;
		}
		LogMessage("INFO", " successfully! [" + inputFilePath + "]");
		LogManager.getLogManager().getMsgPage().log("CALLED " + "\n");
		JOptionPane.showInternalMessageDialog(getContentPane(), ExtensionResources.format("SUCCESS_IMPORT"),
				ExtensionResources.format("SUCCESS_TITLE"), JOptionPane.INFORMATION_MESSAGE);

	}

	public Connection getConnection(String connectionName) throws DBException {
		return Connections.getInstance().getConnection(connectionName);
	}

	private static final void LogMessage(String level, String msg) {
		DateFormat dateFormat = new SimpleDateFormat("HH:mm:ss");
		Date date = new Date();
		String currentTime = "[" + dateFormat.format(date) + "] ";
		LogManager.getLogManager().getMsgPage().log("IMPORT " + currentTime + level + ": " + msg + "\n");
	}

	class ChkItemAdapter implements ItemListener {
		@Override
		public void itemStateChanged(ItemEvent itemE) {
			JRadioButton chkSelected = (JRadioButton) itemE.getSource();
			if (chkSelected != null) {
				if (ExtensionResources.format("DELETE_BEFORE_IMPORT").equals(chkSelected.getText())) {
					deleteBeforeImport = true;
				}
				if (ExtensionResources.format("NOT_DELETE_BEFORE_IMPORT").equals(chkSelected.getText())) {
					deleteBeforeImport = false;
				}
			}
		}
	}
}