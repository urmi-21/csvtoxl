package csvtoxl;

import java.awt.EventQueue;

import javax.swing.JFrame;
import java.awt.BorderLayout;
import javax.swing.JPanel;

import com.univocity.parsers.csv.CsvFormat;
import com.univocity.parsers.csv.CsvParser;
import com.univocity.parsers.csv.CsvParserSettings;

import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JButton;
import javax.swing.JFileChooser;

import java.awt.event.ActionListener;
import java.io.File;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.awt.event.ActionEvent;

public class ReadFilesFrame {

	private JFrame frame;
	private List<File> selectedfiles = new ArrayList<>();
	JPanel added_files_panel = new JPanel();

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

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() {
		frame = new JFrame();
		frame.setBounds(100, 100, 450, 300);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frame.getContentPane().setLayout(new BorderLayout(0, 0));
		
		JPanel panel = new JPanel();
		frame.getContentPane().add(panel, BorderLayout.NORTH);
		
		JLabel lblBrowseFiles = new JLabel("Browse files");
		panel.add(lblBrowseFiles);
		
		JButton btnSelectFiles = new JButton("Select files...");
		btnSelectFiles.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				//browse files
				JFileChooser chooser = new JFileChooser();
				chooser.setMultiSelectionEnabled(true);
				chooser.showOpenDialog(frame);
				File[] files = chooser.getSelectedFiles();
				//add to selected files list
				selectedfiles.addAll(Arrays.asList(files));
				//update display
				updateAddedFiles(files);
				frame.revalidate();
				frame.repaint();
			}
		});
		panel.add(btnSelectFiles);
		
		JPanel panel_1 = new JPanel();
		frame.getContentPane().add(panel_1, BorderLayout.SOUTH);
		
		JLabel lblConvert = new JLabel("Convert");
		panel_1.add(lblConvert);
		
		JButton btnConvert = new JButton("Convert");
		panel_1.add(btnConvert);
		
		JPanel panel_2 = new JPanel();
		frame.getContentPane().add(panel_2, BorderLayout.CENTER);
		panel_2.setLayout(new BorderLayout(0, 0));
		
		JLabel lblAddedFiles = new JLabel("Added files:");
		panel_2.add(lblAddedFiles, BorderLayout.NORTH);
		
		
		panel_2.add(added_files_panel, BorderLayout.CENTER);
	}
	
	/**
	 * Update display for added files
	 */
	private void updateAddedFiles(File[] files) {
		
		for(File f: files) {
			JOptionPane.showMessageDialog(null, f.getPath());	
			//add this panel to frame panel
			added_files_panel.add(addedFile(f.getPath()));
			
		}
		
		
	}
	
	
	private JPanel addedFile(String fpath) {
		JPanel panel=new JPanel();
		JLabel path=new JLabel(fpath);
		//get delimiter
		CsvParserSettings settings = new CsvParserSettings();
		settings.detectFormatAutomatically();
		CsvParser parser = new CsvParser(settings);
		List<String[]> rows = parser.parseAll(new File("/path/to/your.csv"));
		// if you want to see what it detected
		CsvFormat format = parser.getDetectedFormat();
		JLabel delimiter=new JLabel(format.toString());
		JButton remove=new JButton("Remove");
		
		panel.add(path);
		panel.add(delimiter);
		panel.add(remove);
		
		return panel;
		
	}
	
	
	
	

}
