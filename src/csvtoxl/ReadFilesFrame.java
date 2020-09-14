package csvtoxl;

import java.awt.EventQueue;

import javax.swing.JFrame;
import java.awt.BorderLayout;
import java.awt.Color;

import javax.swing.JPanel;
import javax.swing.JProgressBar;
import javax.swing.JScrollPane;
import javax.swing.JSeparator;
import javax.swing.ScrollPaneConstants;
import javax.swing.SwingUtilities;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.univocity.parsers.csv.CsvFormat;
import com.univocity.parsers.csv.CsvParser;
import com.univocity.parsers.csv.CsvParserSettings;

import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JButton;
import javax.swing.JFileChooser;

import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.TreeSet;
import java.awt.event.ActionEvent;
import javax.swing.BoxLayout;
import javax.swing.JCheckBox;
import java.awt.Font;
import javax.swing.SwingConstants;

public class ReadFilesFrame {

	private JFrame frame;
	private Set<File> selectedfiles = new TreeSet<>();

	JPanel added_files_panel = new JPanel();
	JScrollPane scrollPane = new JScrollPane(added_files_panel, ScrollPaneConstants.VERTICAL_SCROLLBAR_AS_NEEDED,
			ScrollPaneConstants.HORIZONTAL_SCROLLBAR_AS_NEEDED);
	JCheckBox chckbxWriteToSingle = new JCheckBox("Write to single xlsx");
	JButton btnConvert = new JButton("Convert");
	JLabel statusString = new JLabel();

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					ReadFilesFrame window = new ReadFilesFrame();
					window.frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the application.
	 */
	public ReadFilesFrame() {
		initialize();
	}

	public JFrame getFrame() {
		return frame;
	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() {
		frame = new JFrame();
		frame.setBounds(100, 100, 450, 300);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frame.getContentPane().setLayout(new BorderLayout(0, 0));
		frame.setTitle("csvtoxl");

		JPanel panel = new JPanel();
		frame.getContentPane().add(panel, BorderLayout.NORTH);

		JLabel lblBrowseFiles = new JLabel("Add files");
		panel.add(lblBrowseFiles);

		JButton btnSelectFiles = new JButton("Select files...");
		btnSelectFiles.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				// browse files
				JFileChooser chooser = new JFileChooser();
				chooser.setMultiSelectionEnabled(true);
				chooser.showOpenDialog(frame);
				File[] files = chooser.getSelectedFiles();
				// add to selected files list
				selectedfiles.addAll(Arrays.asList(files));
				// update display
				updateAddedFiles();
				frame.revalidate();
				frame.repaint();
			}
		});
		panel.add(btnSelectFiles);

		JButton btnRemoveAllFiles = new JButton("Remove all files");
		btnRemoveAllFiles.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				selectedfiles = new TreeSet<>();
				// update display
				updateAddedFiles();
				frame.revalidate();
				frame.repaint();
			}
		});
		panel.add(btnRemoveAllFiles);

		JSeparator s0 = new JSeparator();
		// set layout as vertical
		s0.setOrientation(SwingConstants.VERTICAL);
		panel.add(s0);
		JSeparator s = new JSeparator();
		// set layout as vertical
		s.setOrientation(SwingConstants.VERTICAL);
		panel.add(s);
		JLabel lblStatus = new JLabel("Status:");
		panel.add(lblStatus);
		statusString = new JLabel("idle");
		statusString.setForeground(Color.BLUE);
		statusString.setFont(new Font("Tahoma", Font.BOLD, 13));
		panel.add(statusString);

		JPanel panel_1 = new JPanel();
		frame.getContentPane().add(panel_1, BorderLayout.SOUTH);

		panel_1.add(chckbxWriteToSingle);

		btnConvert.setBackground(Color.GREEN);
		btnConvert.setOpaque(true);
		btnConvert.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				// read files and convert to xlsx
				convertSelectedFiles(chckbxWriteToSingle.isSelected());
			}
		});
		panel_1.add(btnConvert);

		JPanel panel_2 = new JPanel();
		frame.getContentPane().add(panel_2, BorderLayout.CENTER);
		panel_2.setLayout(new BorderLayout(0, 0));

		JLabel lblAddedFiles = new JLabel("Added files:");
		lblAddedFiles.setHorizontalAlignment(SwingConstants.CENTER);
		lblAddedFiles.setFont(new Font("Tahoma", Font.BOLD, 13));
		panel_2.add(lblAddedFiles, BorderLayout.NORTH);

		// panel_2.add(added_files_panel, BorderLayout.CENTER);
		added_files_panel.setLayout(new BoxLayout(added_files_panel, BoxLayout.Y_AXIS));
		// frameConstraints.gridx = 0;
		// frameConstraints.gridy = 1;
		// frameConstraints.weighty = 1;
		panel_2.add(scrollPane);
	}

	/**
	 * Update display for added files
	 */
	private void updateAddedFiles() {
		added_files_panel.removeAll();
		added_files_panel.revalidate();
		added_files_panel.repaint();
		// sort by name
		if (selectedfiles.size() > 0) {
			for (File f : selectedfiles) {
				// JOptionPane.showMessageDialog(null, f.getPath());
				// add this panel to frame panel
				added_files_panel.add(addFileEntry(f.getPath()));
			}
		}

		frame.revalidate();
		frame.repaint();
	}

	private JPanel addFileEntry(String fpath) {
		JPanel panel = new JPanel();
		JLabel path = new JLabel(fpath);
		// get delimiter
		CsvParserSettings settings = new CsvParserSettings();
		settings.detectFormatAutomatically();
		CsvParser parser = new CsvParser(settings);
		List<String[]> rows = parser.parseAll(new File(fpath));
		// if you want to see what it detected
		CsvFormat format = parser.getDetectedFormat();
		String delim = String.valueOf(format.getDelimiter());
		if (delim.equals("\t")) {
			delim = "Tab";
		} else if (delim.equals(" ")) {
			delim = "Space";
		}

		JLabel delimiter = new JLabel(delim);
		JButton remove = new JButton("Remove");
		remove.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				removeSelectedFile(fpath);
			}
		});

		panel.add(path);
		panel.add(delimiter);
		panel.add(remove);

		return panel;

	}

	private void removeSelectedFile(String element) {
		selectedfiles.remove(new File(element));
		updateAddedFiles();
	}

	private void convertSelectedFiles(boolean combine) {
		updateStatus("busy...");
		btnConvert.setEnabled(false);
		btnConvert.setText("Working...");
		btnConvert.setBackground(Color.RED);
		btnConvert.setOpaque(true);
		btnConvert.validate();
		btnConvert.repaint();
		getFrame().validate();
		getFrame().repaint();

		if (combine) {
			JFrame parentFrame = new JFrame();
			JFileChooser fileChooser = new JFileChooser();
			fileChooser.setDialogTitle("Specify out file name");
			int userSelection = fileChooser.showSaveDialog(parentFrame);
			File destination = null;
			if (userSelection == JFileChooser.APPROVE_OPTION) {
				destination = fileChooser.getSelectedFile();
			}
			saveWorkbook(createWorkbook(selectedfiles), destination.getAbsolutePath());

		} else {
			for (File f : selectedfiles) {
				String destination = FilenameUtils.removeExtension(f.getAbsolutePath());
				saveWorkbook(createWorkbook(f), destination);
			}
		}

		// after conversion
		updateStatus("idle");
		btnConvert.setEnabled(true);
		btnConvert.setText("Convert...");
		btnConvert.setBackground(Color.GREEN);
		btnConvert.setOpaque(true);
		btnConvert.validate();
		btnConvert.repaint();
		getFrame().validate();
		getFrame().repaint();

	}

	private boolean saveWorkbook(SXSSFWorkbook workBook, String destination) {
		// write xlsx file
		FileOutputStream out;
		if (!FilenameUtils.getExtension(destination).equals(".xlsx")) {
			destination = destination + ".xlsx";
		}
		try {
			updateStatus("Saving " + destination + " ...");
			out = new FileOutputStream(destination);
			workBook.write(out);
			out.close();
			workBook.close();
			JOptionPane.showMessageDialog(null, "File saved to: " + destination, "File saved",
					JOptionPane.INFORMATION_MESSAGE);
			// dispose of temporary files
			workBook.dispose();
			return true;
		} catch (IOException e) {
			e.printStackTrace();
			JOptionPane.showMessageDialog(null, "Error in saving file: " + destination, "Error",
					JOptionPane.ERROR_MESSAGE);
			// dispose of temporary files
			workBook.dispose();
			return false;
		}
	}

	/**
	 * 
	 * @param f
	 * @return
	 */
	private SXSSFWorkbook createWorkbook(File f) {
		Set<File> fileset = new HashSet<>();
		fileset.add(f);
		return createWorkbook(fileset);
	}

	/**
	 * Save multiple files to single xlsx
	 * 
	 * @param fileset
	 * @return
	 */
	private SXSSFWorkbook createWorkbook(Set<File> fileset) {
		// XSSFWorkbook workBook = new XSSFWorkbook();
		SXSSFWorkbook workBook = new SXSSFWorkbook(1000); // faster

		for (File f : fileset) {
			String sheetName = FilenameUtils.getBaseName(f.getAbsolutePath());
			// update status
			updateStatus("Processing: " + sheetName);
			CsvParserSettings settings = new CsvParserSettings();
			settings.detectFormatAutomatically();
			CsvParser parser = new CsvParser(settings);
			List<String[]> rows = parser.parseAll(f);
			// create sheet
			SXSSFSheet sheet = workBook.createSheet(sheetName);
			SXSSFCell cell = null;
			// set column name colors and font.
			XSSFFont headerFont = (XSSFFont) workBook.createFont();
			headerFont.setColor(IndexedColors.BLACK.index);
			XSSFCellStyle headerCellStyle = (XSSFCellStyle) sheet.getWorkbook().createCellStyle();
			headerCellStyle.setFillForegroundColor(IndexedColors.GOLD.index);
			headerCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			headerCellStyle.setFont(headerFont);

			int rowIndex = 0;
			// XSSFRow row = sheet.createRow(rowIndex);
			for (String[] csvrow : rows) {
				int colIndex = 0;
				SXSSFRow row = sheet.createRow(rowIndex);
				for (String colval : csvrow) {

					cell = row.createCell(colIndex);
					cell.setCellValue(colval);
					if (rowIndex == 0) {
						cell.setCellStyle(headerCellStyle);
					}

					colIndex++;

				}
				rowIndex++;
			}

		}

		return workBook;

	}

	private void updateStatus(String message) {
		statusString.setText(message);
		statusString.validate();
		statusString.repaint();
		getFrame().validate();
		getFrame().repaint();
	}

}

/**
 * Basic progress bar
 * 
 * @author mrbai
 *
 */
class BasicProgressBar extends JPanel {
	JProgressBar pbar;

	public BasicProgressBar(int min, int max) {
		// initialize Progress Bar
		pbar = new JProgressBar();
		pbar.setMinimum(min);
		pbar.setMaximum(max);
		// add to JPanel
		add(pbar);
	}

	public void updateBar(int newValue) {
		pbar.setValue(newValue);
	}
}
