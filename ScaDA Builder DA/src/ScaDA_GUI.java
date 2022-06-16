
/*
*******************************************
* Title:            ScaDA Builder
* Author:           Michael Courteaux
* Email:            mcourte@entergy.com
*******************************************
 */
import java.awt.Color;
import java.awt.EventQueue;
import java.awt.Font;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;

import javax.swing.DefaultListModel;
import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JList;
import javax.swing.JOptionPane;
import javax.swing.JTextField;
import javax.swing.border.LineBorder;

import elements.ComboItem;
import func.Builder;

// TODO: Auto-generated Javadoc
/**
 * The Class ScaDA_GUI.
 */
public class ScaDA_GUI {

	/** The frm scada builder. */
	private JFrame frmScadaBuilder;

	/** The text path. */
	private JTextField textPath;

	/** The path list. */
	DefaultListModel<String> pathList = new DefaultListModel<String>();

	/** The path text. */
	@SuppressWarnings("unused")
	private String pathText = pathList.toString();
	@SuppressWarnings("unused")
	private String feChoice;
	private String chanChoice;
	private String user = System.getProperty("user.name");

	/**
	 * Launch the application.
	 *
	 * @param args the arguments
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			@Override
			public void run() {
				try {
					ScaDA_GUI window = new ScaDA_GUI();
					window.frmScadaBuilder.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the application.
	 */
	public ScaDA_GUI() {
		initialize();
	}

	/**
	 * Initialize the contents of the frame.
	 */
	@SuppressWarnings({ "unchecked", "rawtypes" })
	private void initialize() {
		frmScadaBuilder = new JFrame();
		frmScadaBuilder.setTitle("ScaDA Builder");
		frmScadaBuilder.setBounds(100, 100, 541, 526);
		frmScadaBuilder.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frmScadaBuilder.getContentPane().setLayout(null);

		JButton btnOpen = new JButton("Open");
		btnOpen.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {

				JFileChooser fc = new JFileChooser();
				fc.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
				fc.setDialogTitle("Select Edit Sheet Path");

				if (fc.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
					textPath.setText(fc.getSelectedFile().toString());
				} else {

				}
			}
		});

		JButton btnExec = new JButton("Continue");

		JComboBox feSelect = new JComboBox();
		feSelect.setEnabled(false);
		feSelect.setVisible(false);
		feSelect.setBounds(205, 394, 58, 20);
		frmScadaBuilder.getContentPane().add(feSelect);
		feSelect.addItem(new ComboItem("", ""));
		feSelect.addItem(new ComboItem("FE1", "FE1"));
		feSelect.addItem(new ComboItem("FE2", "FE2"));
		feSelect.addItem(new ComboItem("FE3", "FE3"));
		feSelect.addItem(new ComboItem("FE4", "FE4"));
		feSelect.addItem(new ComboItem("FE5", "FE5"));

		// Select the Front End

		feSelect.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {

				Object fe = feSelect.getSelectedItem();
				feChoice = ((ComboItem) fe).getValue();

			}
		});

		JComboBox chanSelect = new JComboBox();
		chanSelect.setEnabled(false);
		chanSelect.setVisible(false);
		chanSelect.setBounds(174, 425, 89, 20);
		frmScadaBuilder.getContentPane().add(chanSelect);
		chanSelect.addItem(new ComboItem("", ""));
		chanSelect.addItem(new ComboItem("DNP00", "DNP00"));
		chanSelect.addItem(new ComboItem("DNP01", "DNP01"));
		chanSelect.addItem(new ComboItem("DNP02", "DNP02"));
		chanSelect.addItem(new ComboItem("DNP03", "DNP03"));
		chanSelect.addItem(new ComboItem("DNP04", "DNP04"));
		chanSelect.addItem(new ComboItem("DNP05", "DNP05"));
		chanSelect.addItem(new ComboItem("DNP06", "DNP06"));
		chanSelect.addItem(new ComboItem("DNP07", "DNP07"));
		chanSelect.addItem(new ComboItem("DNP08", "DNP08"));
		chanSelect.addItem(new ComboItem("DNP09", "DNP09"));
		chanSelect.addItem(new ComboItem("DNP10", "DNP10"));
		chanSelect.addItem(new ComboItem("DNP11", "DNP11"));
		chanSelect.addItem(new ComboItem("DNP12", "DNP12"));
		chanSelect.addItem(new ComboItem("DNP13", "DNP13"));
		chanSelect.addItem(new ComboItem("DNP14", "DNP14"));
		chanSelect.addItem(new ComboItem("DNP15", "DNP15"));
		chanSelect.addItem(new ComboItem("DNP16", "DNP16"));
		chanSelect.addItem(new ComboItem("DNP17", "DNP17"));
		chanSelect.addItem(new ComboItem("DNP18", "DNP18"));
		chanSelect.addItem(new ComboItem("DNP19", "DNP19"));

		// Select the DNP Channel

		chanSelect.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {

				Object chan = chanSelect.getSelectedItem();
				chanChoice = ((ComboItem) chan).getValue();

				if (chanChoice != "") {
					btnExec.setEnabled(true);
				} else if (chanChoice == "") {
					btnExec.setEnabled(false);
				}
			}
		});

		btnOpen.setBounds(10, 11, 89, 23);
		frmScadaBuilder.getContentPane().add(btnOpen);

		JButton btnLoad = new JButton("Load");
		btnLoad.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {

				pathList.clear();
				File folder = new File(textPath.getText());
				File[] listOfFiles = folder.listFiles();

				for (int i = 0; i < listOfFiles.length; i++) {
					if (listOfFiles[i].toString().endsWith("xlsm")) {

						pathList.addElement(textPath.getText() + "\\" + listOfFiles[i].getName());

					}
				}

			}
		});
		btnLoad.setBounds(426, 11, 89, 23);
		frmScadaBuilder.getContentPane().add(btnLoad);

		textPath = new JTextField();
		textPath.setEditable(false);
		textPath.setBounds(109, 12, 307, 20);
		frmScadaBuilder.getContentPane().add(textPath);
		textPath.setColumns(10);

		JList esList = new JList(pathList);
		esList.setBorder(new LineBorder(new Color(0, 0, 0)));
		esList.setBounds(10, 45, 505, 283);
		frmScadaBuilder.getContentPane().add(esList);

		// Builds the Display Files

		JButton btnDisp = new JButton("Generate Displays");
		btnDisp.addActionListener(new ActionListener() {
			@SuppressWarnings("static-access")
			@Override
			public void actionPerformed(ActionEvent e) {

				Builder b = new Builder();

				for (int i = 0; i < esList.getModel().getSize(); i++) {

					String es = esList.getModel().getElementAt(i).toString();

					try {
						b.readEditSheet(es);
						b.displayGen(b.getRtu(), b.getMenu(), b.getDevType(), b.getDispVer(), b.getLineKv());
					} catch (IOException e1) {
						// TODO Auto-generated catch block
						e1.printStackTrace();
					}

				}

				// Run the compile displays batch file
				// Runtime.getRuntime().exec(null);

				JOptionPane.showMessageDialog(null, "Displays Complete", "Title", 1);
			}
		});

		JLabel lblFE = new JLabel("Front End");
		lblFE.setFont(new Font("Tahoma", Font.PLAIN, 14));
		lblFE.setBounds(288, 394, 61, 17);
		lblFE.setVisible(false);
		frmScadaBuilder.getContentPane().add(lblFE);

		JLabel lblChan = new JLabel("Channel");
		lblChan.setFont(new Font("Tahoma", Font.PLAIN, 14));
		lblChan.setBounds(288, 425, 61, 17);
		lblChan.setVisible(false);
		frmScadaBuilder.getContentPane().add(lblChan);

		btnDisp.setBounds(42, 339, 119, 23);
		frmScadaBuilder.getContentPane().add(btnDisp);

		// Executes the Project Builder and Builds the Project Files

		btnExec.addActionListener(new ActionListener() {
			@SuppressWarnings({ "static-access" })
			@Override
			public void actionPerformed(ActionEvent e) {

				feSelect.setEnabled(false);
				chanSelect.setEnabled(false);
				btnExec.setVisible(false);
				feSelect.setVisible(false);
				chanSelect.setVisible(false);
				lblFE.setVisible(false);
				lblChan.setVisible(false);

				Builder b = new Builder();

				for (int i = 0; i < esList.getModel().getSize(); i++) {

					String es = esList.getModel().getElementAt(i).toString();

					try {
						b.readEditSheet(es);
						b.alarm(b.getRtu());
						b.fgdis(b.getRtu());
						b.comm(b.getRtu(), b.getAor(), feChoice, chanChoice);
					} catch (IOException e1) {
						// TODO Auto-generated catch block
						e1.printStackTrace();
					}

				}

				JOptionPane.showMessageDialog(null, "Build Complete", "Title", 1);
			}
		});

		btnExec.setEnabled(false);
		btnExec.setVisible(false);
		btnExec.setBounds(203, 456, 119, 23);
		frmScadaBuilder.getContentPane().add(btnExec);

		JButton btnBuild = new JButton("Build Project");
		btnBuild.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				feSelect.setEnabled(true);
				chanSelect.setEnabled(true);
				btnExec.setVisible(true);
				feSelect.setVisible(true);
				chanSelect.setVisible(true);
				lblFE.setVisible(true);
				lblChan.setVisible(true);
			}
		});
		btnBuild.setBounds(203, 339, 119, 23);
		frmScadaBuilder.getContentPane().add(btnBuild);

		// Add Linkages to Project

		JButton btnLink = new JButton("Add Linkages");
		btnLink.addActionListener(new ActionListener() {
			@Override
			@SuppressWarnings("static-access")
			public void actionPerformed(ActionEvent e) {

				Builder b = new Builder();

				for (int i = 0; i < esList.getModel().getSize(); i++) {

					String es = esList.getModel().getElementAt(i).toString();

					try {
						b.linker(b.getScadar(), b.getRtu());
					} catch (IOException e1) {
						// TODO Auto-generated catch block
						e1.printStackTrace();
					}

				}

			}
		});
		btnLink.setBounds(364, 339, 119, 23);
		frmScadaBuilder.getContentPane().add(btnLink);

	}
}
