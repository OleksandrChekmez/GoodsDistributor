package com.aurumsoft.util;

import java.awt.Desktop;
import java.awt.Dimension;
import java.awt.EventQueue;
import java.awt.Font;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.Image;
import java.awt.Insets;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.io.Writer;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.Locale;
import java.util.Map.Entry;
import java.util.Properties;
import java.util.Random;
import java.util.Set;
import java.util.StringTokenizer;
import java.util.TreeSet;

import javax.imageio.ImageIO;
import javax.swing.ButtonGroup;
import javax.swing.DefaultListModel;
import javax.swing.GroupLayout;
import javax.swing.GroupLayout.Alignment;
import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JComboBox;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JList;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JProgressBar;
import javax.swing.JRadioButton;
import javax.swing.JScrollPane;
import javax.swing.JSpinner;
import javax.swing.JTabbedPane;
import javax.swing.JTextArea;
import javax.swing.JTextField;
import javax.swing.LayoutStyle.ComponentPlacement;
import javax.swing.SpinnerDateModel;
import javax.swing.SpinnerNumberModel;
import javax.swing.UIManager;
import javax.swing.filechooser.FileFilter;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class MainWindow {

	private static final Log log = LogFactory.getLog(MainWindow.class);
	private static final String dayMoneyMinProp = "dayMoneyMin";
	private static final String dayMoneyMaxProp = "dayMoneyMax";
	private static final String spinnerMaxGoodsQuantityProp = "spinnerMaxGoodsQuantity";
	private static final String checkBoxOpenFilesProp = "checkBoxOpenFiles";
	private static final String checkBoxDoNotChangeWarehouseProp = "checkBoxDoNotChangeWarehouse";

	private static final String iconName = "coins.png";
	private static Image iconImage;
	private JFrame frame;
	private JTextField warehouseFilePathField;
	private JTextField statementFilePathField;
	private JButton btnCalculate;
	private JSpinner spinnerMoneyTo;
	private JSpinner spinnerMoneyFrom;
	private JSpinner spinnerYear;
	private JCheckBox checkBoxOpenFiles;
	private JCheckBox checkBoxDoNotChangeWarehouse;
	private JButton btnSelectWarehouseFile;
	private JButton btnSelectStatementFile;
	private static String prevDir = null;
	private JTextArea excludedGoodsListArea;
	private static final int totalStatementRow = 79;
	private JComboBox<String> comboBoxMonth;
	private JRadioButton radioButtonFirst;
	private JRadioButton radioButtonSecond;
	private JRadioButton radioButtonFull;
	private JProgressBar progressBar;
	private Properties props = new Properties();

	// FEFF because this is the Unicode char represented by the UTF-8 byte order
	// mark (EF BB BF).
	public static final String UTF8_BOM = "\uFEFF";

	private JSpinner spinnerMaxGoodsQuantity;

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		try {
			UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
		} catch (Exception ex) {
		}
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					Locale.setDefault(Locale.forLanguageTag("ru"));
					MainWindow window = new MainWindow();
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
	public MainWindow() {
		initialize();
	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() {
		frame = new JFrame();

		frame.setTitle("Товарный калькулятор v1.6.1");
		frame.setIconImage(getIcon());
		frame.setBounds(100, 100, 790, 320);
		frame.setMinimumSize(new Dimension(790, 320));
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

		JTabbedPane tabbedPane = new JTabbedPane();
		JPanel container1 = new JPanel();
		tabbedPane.addTab("Набивачка", container1);

		JPanel container2 = new JPanel();
		tabbedPane.addTab("Сума", container2);

		createSumPane(container2);

		JLabel label = new JLabel("Остатки:");

		warehouseFilePathField = new JTextField();
		warehouseFilePathField.setColumns(10);

		btnSelectWarehouseFile = new JButton("Выбрать...");
		btnSelectWarehouseFile.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				File[] f = selectFolder(false);
				if (f != null)
					warehouseFilePathField.setText(f[0].getAbsolutePath());
			}
		});

		JLabel label_1 = new JLabel("Ведомость реализации:");

		statementFilePathField = new JTextField();
		statementFilePathField.setColumns(10);

		btnSelectStatementFile = new JButton("Выбрать...");
		btnSelectStatementFile.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				File[] f = selectFolder(false);
				if (f != null)
					statementFilePathField.setText(f[0].getAbsolutePath());
			}
		});

		JLabel label_2 = new JLabel("Период для рассчетов:");

		Calendar cal = Calendar.getInstance(new Locale("ru"));
		Date initDate = cal.getTime();
		cal.add(Calendar.YEAR, -100);
		Date earliestDate = cal.getTime();
		cal.add(Calendar.YEAR, 200);
		Date latestDate = cal.getTime();
		SpinnerDateModel dateModel = new SpinnerDateModel(initDate, earliestDate, latestDate, Calendar.YEAR);
		spinnerYear = new JSpinner(dateModel);
		spinnerYear.setEditor(new JSpinner.DateEditor(spinnerYear, "yyyy"));

		SpinnerNumberModel moneyModel1 = new SpinnerNumberModel(200, 1, 100000, 1);
		spinnerMoneyFrom = new JSpinner(moneyModel1);

		JLabel label_4 = new JLabel("-");

		SpinnerNumberModel moneyModel2 = new SpinnerNumberModel(300, 1, 100000, 1);
		spinnerMoneyTo = new JSpinner(moneyModel2);

		JLabel label_5 = new JLabel("Разрешенный диапазон выручки за день:");

		btnCalculate = new JButton("Распределить товары");
		btnCalculate.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				Thread t = new Thread(new Runnable() {

					@Override
					public void run() {
						enableControls(false);
						progressBar.setIndeterminate(true);
						try {
							boolean isOpenExcel = checkBoxOpenFiles.isSelected();
							if (radioButtonFirst.isSelected()) {
								calculate(true, isOpenExcel);
							} else if (radioButtonSecond.isSelected()) {
								calculate(false, isOpenExcel);
							} else {
								calculate(true, false);
								calculate(false, isOpenExcel);
							}
						} catch (Exception e) {
							log.error(e.getLocalizedMessage(), e);
							JOptionPane.showMessageDialog(frame,
									"Не предвиденная ошибка:" + e.getClass().getSimpleName() + " - "
											+ e.getLocalizedMessage() + "\n" + getStackTrace(e),
									"Ошибка", JOptionPane.ERROR_MESSAGE);
						}
						progressBar.setIndeterminate(false);
						enableControls(true);
					}
				});
				t.start();
			}
		});
		btnCalculate.setFont(new Font("Tahoma", Font.BOLD, 11));

		checkBoxOpenFiles = new JCheckBox("Открыть файлы после расчетов");
		checkBoxOpenFiles.setSelected(true);

		checkBoxDoNotChangeWarehouse = new JCheckBox("Не менять файл остатков");

		JLabel label_6 = new JLabel("Группы товаров:");

		JScrollPane scrollPane = new JScrollPane();

		comboBoxMonth = new JComboBox<>();

		for (int i = 0; i < 12; i++) {
			cal.set(Calendar.MONTH, i);
			comboBoxMonth.addItem(cal.getDisplayName(Calendar.MONTH, Calendar.LONG, new Locale("ru")));
		}

		ButtonGroup bg = new ButtonGroup();
		radioButtonFirst = new JRadioButton("первая половина");
		radioButtonSecond = new JRadioButton("вторая половина");
		radioButtonFull = new JRadioButton("весь месяц");
		bg.add(radioButtonFirst);
		bg.add(radioButtonSecond);
		bg.add(radioButtonFull);

		cal = Calendar.getInstance(new Locale("ru"));
		int currentMonth = cal.get(Calendar.MONTH);
		int today = cal.get(Calendar.DAY_OF_MONTH);
		if (today > 15 && currentMonth < 11) {
			comboBoxMonth.setSelectedIndex(currentMonth + 1);
			radioButtonFirst.setSelected(true);
		} else {
			comboBoxMonth.setSelectedIndex(currentMonth);
			radioButtonSecond.setSelected(true);
		}

		progressBar = new JProgressBar();

		SpinnerNumberModel quantModel = new SpinnerNumberModel(28, 10, 75, 1);
		spinnerMaxGoodsQuantity = new JSpinner(quantModel);

		JLabel label_3 = new JLabel("Максимальное количество товара:");

		GroupLayout groupLayout = new GroupLayout(container1);
		groupLayout
				.setHorizontalGroup(groupLayout.createParallelGroup(Alignment.TRAILING)
						.addGroup(groupLayout.createSequentialGroup().addContainerGap()
								.addGroup(groupLayout
										.createParallelGroup(Alignment.TRAILING).addGroup(groupLayout
												.createSequentialGroup().addGroup(
														groupLayout.createParallelGroup(Alignment.TRAILING)
																.addComponent(label_3).addComponent(label_2)
																.addComponent(label_1).addComponent(label).addComponent(
																		label_5)
																.addComponent(label_6))
												.addPreferredGap(ComponentPlacement.RELATED).addGroup(
														groupLayout.createParallelGroup(Alignment.LEADING)
																.addGroup(groupLayout.createSequentialGroup()
																		.addComponent(checkBoxOpenFiles)
																		.addPreferredGap(ComponentPlacement.RELATED,
																				101, Short.MAX_VALUE)
																		.addComponent(btnCalculate))
																.addGroup(groupLayout
																		.createSequentialGroup().addGroup(groupLayout
																				.createParallelGroup(Alignment.LEADING)
																				.addComponent(statementFilePathField,
																						GroupLayout.DEFAULT_SIZE, 362,
																						Short.MAX_VALUE)
																				.addComponent(warehouseFilePathField,
																						GroupLayout.DEFAULT_SIZE, 362,
																						Short.MAX_VALUE))
																		.addPreferredGap(ComponentPlacement.RELATED)
																		.addGroup(groupLayout.createParallelGroup(
																				Alignment.LEADING).addComponent(
																						btnSelectWarehouseFile)
																				.addComponent(btnSelectStatementFile)))
																.addComponent(checkBoxDoNotChangeWarehouse)
																.addComponent(
																		scrollPane, GroupLayout.DEFAULT_SIZE, 457,
																		Short.MAX_VALUE)
																.addGroup(groupLayout.createSequentialGroup()
																		.addGroup(groupLayout
																				.createParallelGroup(Alignment.TRAILING,
																						false)
																				.addComponent(spinnerMaxGoodsQuantity,
																						Alignment.LEADING)
																				.addComponent(spinnerMoneyFrom,
																						Alignment.LEADING, 0, 0,
																						Short.MAX_VALUE)
																				.addComponent(
																						spinnerYear, Alignment.LEADING,
																						GroupLayout.PREFERRED_SIZE, 54,
																						Short.MAX_VALUE))
																		.addPreferredGap(ComponentPlacement.RELATED)
																		.addComponent(label_4)
																		.addPreferredGap(ComponentPlacement.RELATED)
																		.addGroup(groupLayout
																				.createParallelGroup(Alignment.LEADING)
																				.addComponent(spinnerMoneyTo,
																						GroupLayout.PREFERRED_SIZE, 54,
																						GroupLayout.PREFERRED_SIZE)
																				.addGroup(groupLayout
																						.createSequentialGroup()
																						.addComponent(comboBoxMonth,
																								GroupLayout.PREFERRED_SIZE,
																								93,
																								GroupLayout.PREFERRED_SIZE)
																						.addPreferredGap(
																								ComponentPlacement.RELATED)
																						.addComponent(radioButtonFirst)
																						.addPreferredGap(
																								ComponentPlacement.RELATED)
																						.addComponent(radioButtonSecond)
																						.addPreferredGap(
																								ComponentPlacement.RELATED)
																						.addComponent(radioButtonFull)))

																		.addGap(67))))
										.addComponent(progressBar, GroupLayout.DEFAULT_SIZE, 674, Short.MAX_VALUE))
								.addContainerGap()));
		groupLayout.setVerticalGroup(groupLayout.createParallelGroup(Alignment.LEADING).addGroup(groupLayout
				.createSequentialGroup().addContainerGap()
				.addGroup(groupLayout.createParallelGroup(Alignment.BASELINE).addComponent(label)
						.addComponent(warehouseFilePathField, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE,
								GroupLayout.PREFERRED_SIZE)
						.addComponent(btnSelectWarehouseFile))
				.addPreferredGap(ComponentPlacement.RELATED)
				.addGroup(groupLayout.createParallelGroup(Alignment.BASELINE).addComponent(label_1)
						.addComponent(statementFilePathField, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE,
								GroupLayout.PREFERRED_SIZE)
						.addComponent(btnSelectStatementFile))
				.addPreferredGap(ComponentPlacement.RELATED)
				.addGroup(groupLayout.createParallelGroup(Alignment.BASELINE).addComponent(label_2)
						.addComponent(spinnerYear, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE,
								GroupLayout.PREFERRED_SIZE)
						.addComponent(comboBoxMonth, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE,
								GroupLayout.PREFERRED_SIZE)
						.addComponent(radioButtonFirst).addComponent(radioButtonSecond).addComponent(radioButtonFull))
				.addPreferredGap(ComponentPlacement.RELATED)
				.addGroup(groupLayout.createParallelGroup(Alignment.BASELINE)
						.addComponent(spinnerMoneyFrom, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE,
								GroupLayout.PREFERRED_SIZE)
						.addComponent(label_5).addComponent(label_4).addComponent(spinnerMoneyTo,
								GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE))
				.addPreferredGap(ComponentPlacement.RELATED)
				.addGroup(groupLayout.createParallelGroup(Alignment.BASELINE)
						.addComponent(spinnerMaxGoodsQuantity, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE,
								GroupLayout.PREFERRED_SIZE)
						.addComponent(label_3))
				.addPreferredGap(ComponentPlacement.RELATED)
				.addGroup(groupLayout.createParallelGroup(Alignment.LEADING).addComponent(label_6)
						.addComponent(scrollPane, GroupLayout.DEFAULT_SIZE, 51, Short.MAX_VALUE))
				.addPreferredGap(ComponentPlacement.RELATED)
				.addGroup(groupLayout.createParallelGroup(Alignment.TRAILING)
						.addGroup(groupLayout.createSequentialGroup().addComponent(checkBoxDoNotChangeWarehouse)
								.addPreferredGap(ComponentPlacement.RELATED).addComponent(checkBoxOpenFiles).addGap(10))
						.addGroup(groupLayout.createSequentialGroup().addComponent(btnCalculate)
								.addPreferredGap(ComponentPlacement.RELATED)))
				.addComponent(progressBar, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE,
						GroupLayout.PREFERRED_SIZE)
				.addContainerGap()));

		excludedGoodsListArea = new JTextArea();
		excludedGoodsListArea.setLineWrap(true);
		excludedGoodsListArea.setFont(new Font("Tahoma", Font.PLAIN, 11));
		scrollPane.setViewportView(excludedGoodsListArea);
		loadSavedOptions();
		container1.setLayout(groupLayout);
		frame.getContentPane().add(tabbedPane);

	}

	private void createSumPane(JPanel container2) {
		JLabel label = new JLabel("Файлы для суммирования:");

		DefaultListModel<File> selectedFilesToAddModel = new DefaultListModel<File>();
		JList<File> selectedFilesToAdd = new JList<File>(selectedFilesToAddModel);
		JScrollPane listScroller = new JScrollPane(selectedFilesToAdd);
		JButton btnAddFile = new JButton("Добавить...");
		btnAddFile.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				File[] f = selectFolder(true);
				if (f != null) {
					for (File file : f) {
						selectedFilesToAddModel.addElement(file);
					}

				}
			}
		});
		JButton btnRemoveFile = new JButton("Удалить");
		btnRemoveFile.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				File f = selectedFilesToAdd.getSelectedValue();
				if (f != null) {
					selectedFilesToAddModel.removeElement(f);
				}
			}
		});

		JButton btnRemoveAllFile = new JButton("Удалить все");
		btnRemoveAllFile.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				selectedFilesToAddModel.removeAllElements();

			}
		});
		JButton btnSumm = new JButton("Cуммировать");
		btnSumm.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				sumFiles(selectedFilesToAddModel);
			}
		});
		container2.setLayout(new GridBagLayout());
		GridBagConstraints c = new GridBagConstraints();
		c.gridx = 0;
		c.gridy = 0;
		c.anchor = GridBagConstraints.WEST;
		c.insets = new Insets(2, 2, 2, 2);
		container2.add(label, c);
		c.gridx = 0;
		c.gridy = 1;
		c.weightx = 1;
		c.weighty = 1;
		c.fill = GridBagConstraints.BOTH;
		c.gridheight = 4;
		container2.add(listScroller, c);
		c.gridx = 1;
		c.weightx = 0;
		c.weighty = 0;
		c.gridheight = 1;
		c.gridy = 1;
		c.fill = GridBagConstraints.HORIZONTAL;
		container2.add(btnAddFile, c);
		c.gridx = 1;
		c.gridy = 2;
		container2.add(btnRemoveFile, c);
		c.gridx = 1;
		c.gridy = 3;
		container2.add(btnRemoveAllFile, c);
		c.gridx = 1;
		c.gridy = 4;
		c.anchor = GridBagConstraints.NORTHWEST;
		container2.add(btnSumm, c);
	}

	protected void sumFiles(DefaultListModel<File> selectedFilesToAddModel) {
		String targetFile = selectSaveFile();
		if (targetFile != null) {
			HashMap<String, SoldGoods> soldsGoods = new HashMap<String, SoldGoods>();
			String fopName = "";
			for (int i = 0; i < selectedFilesToAddModel.getSize(); i++) {
				File file = selectedFilesToAddModel.get(i);
				String fopNameTmp = readSourceFile(file, soldsGoods);
				if (fopNameTmp != null) {
					fopName = fopNameTmp;
				}
			}

			saveGoodsTotal(soldsGoods, fopName, targetFile);
		}
	}

	private void saveGoodsTotal(HashMap<String, SoldGoods> soldsGoods, String fopName, String targetFile) {
		InputStream fileStream = null;
		String templateFile = "sum_template.xls";
		try {
			fileStream = new FileInputStream(templateFile);
		} catch (FileNotFoundException e) {
			log.error(e.getLocalizedMessage(), e);
			JOptionPane.showMessageDialog(frame, "Файл шаблона не найден: " + templateFile, "Ошибка",
					JOptionPane.ERROR_MESSAGE);
			return;
		}
		Workbook soldGoodsWorkBook = null;
		try {
			soldGoodsWorkBook = WorkbookFactory.create(fileStream);
		} catch (Exception e) {
			log.error(e.getLocalizedMessage(), e);
			JOptionPane.showMessageDialog(frame,
					"Невозможно открыть файл шаблона - " + templateFile
							+ "\\nВозможно он поврежден.\nОткройте его в MS Excel и пересохраните!" + "\nОшибка: "
							+ e.getLocalizedMessage(),
					"Ошибка", JOptionPane.ERROR_MESSAGE);
			return;
		}
		Sheet sheet1 = soldGoodsWorkBook.getSheetAt(0);
		Sheet sheet2 = soldGoodsWorkBook.getSheetAt(1);
		writeSoldGoodsSheet(sheet1, soldsGoods, fopName);
		writeSoldGoodsSheet(sheet2, soldsGoods, fopName);

		try {

			// save workbook
			log.debug("save workbook");
			FileOutputStream fileOut = new FileOutputStream(targetFile);
			soldGoodsWorkBook.write(fileOut);
			fileOut.close();
			log.debug("save done");

			try {
				Desktop desktop = null;
				if (Desktop.isDesktopSupported()) {
					desktop = Desktop.getDesktop();
				}
				if (desktop != null) {
					desktop.open(new File(targetFile));
				}
			} catch (IOException ioe) {
				log.error("Cannot open report in default viewer", ioe);
			}

		} catch (Exception e) {
			log.error(e.getLocalizedMessage(), e);
			JOptionPane
					.showMessageDialog(frame,
							"Ошибка при сохранении файла результатов!\n" + e.getClass().getSimpleName() + " - "
									+ e.getLocalizedMessage() + "\n" + getStackTrace(e),
							"Ошибка", JOptionPane.ERROR_MESSAGE);
		}
	}

	private void writeSoldGoodsSheet(Sheet sheet1, HashMap<String, SoldGoods> soldsGoods, String fopName) {
		try {
			// write FOP name
			Row topRow = sheet1.getRow(0);
			for (Cell c : topRow) {
				String cVal = getCellValue(c);
				if (cVal != null && "ФОП".equals(cVal)) {
					c.setCellValue(fopName);
				}
			}

			// save goods
			int headerRow = -1;
			HashMap<Integer, Integer> dayColumnMap = new HashMap<Integer, Integer>();
			ArrayList<Integer> nameColumns = new ArrayList<Integer>();
			int totalColumn = -1;
			for (Row row : sheet1) {
				// find table header by "Наименование"
				Cell firstColumnCell = row.getCell(0);
				String cellVal = getCellValue(firstColumnCell);
				if (cellVal != null) {
					if (headerRow == -1) {
						if ("Наименование".equals(cellVal)) {
							headerRow = row.getRowNum();
							// read header columns
							Row dayNumberRow = sheet1.getRow(headerRow - 1);
							for (int i = 0; i < dayNumberRow.getLastCellNum(); i++) {
								Cell dCell = dayNumberRow.getCell(i);
								Cell headCell = row.getCell(i);
								String dayVal = getCellValue(dCell);
								String headerVal = getCellValue(headCell);

								if ("Наименование".equals(headerVal)) {
									nameColumns.add(i);
								}

								if (dayVal != null && "Кол-во тов".equals(headerVal)) {
									dayColumnMap.put(Integer.parseInt(dayVal), i);
								}

							}
						}
					} else {
						break;
					}
				}
			}
			// write goods list
			int goodsRow = headerRow + 1;
			SoldGoods totalRow = new SoldGoods();
			totalRow.setName("Итого");
			for (Entry<String, SoldGoods> en : soldsGoods.entrySet()) {
				String goodsName = en.getKey();
				SoldGoods goodsData = en.getValue();
				Row r = sheet1.getRow(goodsRow);

				// write name and sell price
				for (Integer nameColumn : nameColumns) {
					Cell c = r.getCell(nameColumn);
					c.setCellValue(goodsName);
					c = r.getCell(nameColumn + 1);
					c.setCellValue(goodsData.getSellPrice());
				}

				// write data by days
				for (Entry<Integer, Integer> den : dayColumnMap.entrySet()) {
					Integer day = den.getKey();
					Integer column = den.getValue();
					// Кол-во тов
					Cell c = r.getCell(column);
					Integer qty = goodsData.getQuantities().get(day);
					if (qty != null) {
						c.setCellValue(qty);
					} else {
						c.setCellValue(0.0);
					}
					// Сумма реал.с НДС
					c = r.getCell(column + 1);
					Double realPrice = goodsData.getPriceRealNDS().get(day);
					Double totalRealPrice = totalRow.getPriceRealNDS().get(day);
					if (realPrice != null) {
						c.setCellValue(realPrice);
						if (totalRealPrice != null) {
							totalRow.getPriceRealNDS().put(day, totalRealPrice + realPrice);
						} else {
							totalRow.getPriceRealNDS().put(day, realPrice);
						}
					} else {
						c.setCellValue(0.0);
					}

					// Сумма НДС
					c = r.getCell(column + 2);
					Double price = goodsData.getPriceNDS().get(day);
					Double totalPrice = totalRow.getPriceNDS().get(day);
					if (price != null) {
						c.setCellValue(price);
						if (totalPrice != null) {
							totalRow.getPriceNDS().put(day, price + totalPrice);
						} else {
							totalRow.getPriceNDS().put(day, price);
						}
					} else {
						c.setCellValue(0.0);
					}

					// Сумма без НДС
					c = r.getCell(column + 3);
					if (realPrice != null) {
						c.setCellValue(realPrice / 1.2);
					} else {
						c.setCellValue(0.0);
					}

					// Цена без НДС
					c = r.getCell(column + 4);
					if (realPrice != null && qty != null && qty > 0) {
						c.setCellValue((realPrice / 1.2) / qty);
					} else {
						c.setCellValue(0.0);
					}
				}

				goodsRow++;
			}

			// write total row
			Row r = sheet1.getRow(goodsRow);
			Cell cell = r.getCell(0);
			Workbook wb = cell.getRow().getSheet().getWorkbook();
			CellStyle style = cell.getCellStyle();
			org.apache.poi.ss.usermodel.Font oldFont = wb.getFontAt(style.getFontIndex());

			org.apache.poi.ss.usermodel.Font boldFont = wb.createFont();
			boldFont.setBold(true);
			boldFont.setCharSet(oldFont.getCharSet());
			boldFont.setFontHeightInPoints(oldFont.getFontHeightInPoints());
			boldFont.setItalic(oldFont.getItalic());
			boldFont.setStrikeout(oldFont.getStrikeout());
			boldFont.setTypeOffset(oldFont.getTypeOffset());
			boldFont.setUnderline(oldFont.getUnderline());
			boldFont.setColor(oldFont.getColor());
			// write name and sell price
			for (Integer nameColumn : nameColumns) {
				Cell c = r.getCell(nameColumn);
				c.setCellValue(totalRow.getName());
				CellUtil.setCellStyleProperty(c, CellUtil.FONT, boldFont.getIndex());
			}
			// write data by days
			for (Entry<Integer, Integer> den : dayColumnMap.entrySet()) {
				Integer day = den.getKey();
				Integer column = den.getValue();

				// Сумма реал.с НДС
				Cell c = r.getCell(column + 1);
				CellUtil.setCellStyleProperty(c, CellUtil.FONT, boldFont.getIndex());
				Double realPrice = totalRow.getPriceRealNDS().get(day);
				if (realPrice != null) {
					c.setCellValue(realPrice);
				} else {
					c.setCellValue(0.0);
				}
				// save Сумма реал.с НДС into Quantity total
				c = r.getCell(column);
				CellUtil.setCellStyleProperty(c, CellUtil.FONT, boldFont.getIndex());
				if (realPrice != null) {
					c.setCellValue(realPrice);
				} else {
					c.setCellValue(0.0);
				}

				// Сумма НДС
				c = r.getCell(column + 2);
				CellUtil.setCellStyleProperty(c, CellUtil.FONT, boldFont.getIndex());
				Double price = totalRow.getPriceNDS().get(day);
				if (price != null) {
					c.setCellValue(price);
				} else {
					c.setCellValue(0.0);
				}

				// Сумма без НДС
				c = r.getCell(column + 3);
				CellUtil.setCellStyleProperty(c, CellUtil.FONT, boldFont.getIndex());
				if (realPrice != null) {
					c.setCellValue(realPrice / 1.2);
				} else {
					c.setCellValue(0.0);
				}
			}

			// hide sell price column and names except first
			for (Integer nameColumn : nameColumns) {
				if (nameColumn != 0)
					sheet1.setColumnHidden(nameColumn, true);
				sheet1.setColumnHidden(nameColumn + 1, true);
			}
			// hide columns with NDS prices
			for (Entry<Integer, Integer> den : dayColumnMap.entrySet()) {
				Integer column = den.getValue();
				sheet1.setColumnHidden(column + 1, true);
				sheet1.setColumnHidden(column + 2, true);
				sheet1.setColumnWidth(column, 1950);
				sheet1.setColumnWidth(column + 3, 1950);
//				sheet1.setColumnWidth(column + 4, 1960);
			}
		} catch (Exception e) {
			log.error(e.getLocalizedMessage(), e);
			JOptionPane
					.showMessageDialog(frame,
							"Ошибка записи данных в результирующий файл!\n" + e.getClass().getSimpleName() + " - "
									+ e.getLocalizedMessage() + "\n" + getStackTrace(e),
							"Ошибка", JOptionPane.ERROR_MESSAGE);
			return;
		}
	}

	private String readSourceFile(File file, HashMap<String, SoldGoods> soldsGoods) {
		InputStream fileStream = null;
		String fopName = null;
		try {
			fileStream = new FileInputStream(file);
		} catch (FileNotFoundException e) {
			log.error(e.getLocalizedMessage(), e);
			JOptionPane.showMessageDialog(frame, "Файл не найден: " + file.getAbsolutePath(), "Ошибка",
					JOptionPane.ERROR_MESSAGE);
			return null;
		}
		Workbook soldGoodsWorkBook = null;
		try {
			soldGoodsWorkBook = WorkbookFactory.create(fileStream);
		} catch (Exception e) {
			log.error(e.getLocalizedMessage(), e);
			JOptionPane.showMessageDialog(frame,
					"Невозможно открыть файл - " + file.getName()
							+ "\\nВозможно он поврежден.\nОткройте его в MS Excel и пересохраните!" + "\nОшибка: "
							+ e.getLocalizedMessage(),
					"Ошибка", JOptionPane.ERROR_MESSAGE);
			return null;
		}
		Sheet sheet1 = soldGoodsWorkBook.getSheetAt(0);
		Sheet sheet2 = soldGoodsWorkBook.getSheetAt(1);
		readSoldGoodsSheet(file, sheet1, soldsGoods);
		readSoldGoodsSheet(file, sheet2, soldsGoods);
		try {
			fopName = getCellValue(sheet1.getRow(0).getCell(0));
		} catch (Exception e) {
			log.error(e.getMessage(), e);
		}
		return fopName;

	}

	private void readSoldGoodsSheet(File file, Sheet sheet1, HashMap<String, SoldGoods> soldsGoods) {
		int headerRow = -1;
		HashMap<Integer, Integer> dayColumnMap = new HashMap<Integer, Integer>();
		try {
			for (Row row : sheet1) {
				// find table header by "Наименование"
				Cell firstColumnCell = row.getCell(0);
				String cellVal = getCellValue(firstColumnCell);
				if (cellVal != null) {
					if (headerRow == -1) {
						if ("Наименование".equals(cellVal)) {
							headerRow = row.getRowNum();
							// read header columns
							Row dayNumberRow = sheet1.getRow(headerRow - 1);
							for (int i = 0; i < dayNumberRow.getLastCellNum(); i++) {
								Cell dCell = dayNumberRow.getCell(i);
								Cell headCell = row.getCell(i);
								String dayVal = getCellValue(dCell);
								String headerVal = getCellValue(headCell);
								if ("Итого".equals(dayVal)) {
									break;
								}
								if (dayVal != null && "Кол-во тов".equals(headerVal)) {
									dayColumnMap.put(Integer.parseInt(dayVal), i);
								}
							}
						}
					} else {
						// read goods list
						if ("0.0".equals(cellVal)) {
							break;
						}
						SoldGoods g = soldsGoods.get(cellVal);
						if (g == null) {
							g = new SoldGoods();
							g.setName(cellVal);

							Cell priceCell = row.getCell(1);
							String sellPriceVal = getCellValue(priceCell);
							try {
								g.setSellPrice(Double.parseDouble(sellPriceVal));
							} catch (Exception e) {
								log.error(e.getLocalizedMessage(), e);
								JOptionPane.showMessageDialog(frame,
										"Продажная цена не цифра, строка " + (row.getRowNum() + 1) + ", лист "
												+ sheet1.getSheetName() + " в файле " + file.getName() + "!\n"
												+ e.getClass().getSimpleName() + " - " + e.getLocalizedMessage() + "\n"
												+ getStackTrace(e),
										"Ошибка", JOptionPane.ERROR_MESSAGE);
								return;

							}
							soldsGoods.put(g.getName(), g);
						}
						// read quantities
						for (Entry<Integer, Integer> en : dayColumnMap.entrySet()) {
							Integer day = en.getKey();
							Integer column = en.getValue();
							Cell qtyCell = row.getCell(column);
							Cell priceRealCell = row.getCell(column + 1);
							Cell priceCell = row.getCell(column + 2);
							String qtyVal = getCellValue(qtyCell);
							String priceRealVal = getCellValue(priceRealCell);
							String priceVal = getCellValue(priceCell);
							if (qtyVal != null) {
								int qty = (int) Math.round(Double.parseDouble(qtyVal));
								double priceReal = Double.parseDouble(priceRealVal);
								double price = Double.parseDouble(priceVal);
								if (qty > 0) {
									Integer oldQty = g.getQuantities().get(day);
									if (oldQty == null) {
										g.getQuantities().put(day, qty);
									} else {
										g.getQuantities().put(day, qty + oldQty);
									}
									Double oldRealPrice = g.getPriceRealNDS().get(day);
									if (oldRealPrice == null) {
										g.getPriceRealNDS().put(day, priceReal);
									} else {
										g.getPriceRealNDS().put(day, priceReal + oldRealPrice);
									}
									Double oldPrice = g.getPriceNDS().get(day);
									if (oldPrice == null) {
										g.getPriceNDS().put(day, price);
									} else {
										g.getPriceNDS().put(day, price + oldPrice);
									}
								}
							}
						}

					}
				}
			}

		} catch (Exception e) {
			log.error(e.getLocalizedMessage(), e);
			JOptionPane
					.showMessageDialog(frame,
							"Ошибка при работе с файлом " + file.getName() + "!\n" + e.getClass().getSimpleName()
									+ " - " + e.getLocalizedMessage() + "\n" + getStackTrace(e),
							"Ошибка", JOptionPane.ERROR_MESSAGE);
			return;
		}
	}

	private static String removeUTF8BOM(String s) {
		if (s.startsWith(UTF8_BOM)) {
			s = s.substring(1);
		}
		return s;
	}

	private void loadSavedOptions() {
		try (BufferedReader br = new BufferedReader(
				new InputStreamReader(new FileInputStream("excluded.txt"), "UTF-8"))) {
			StringBuilder sb = new StringBuilder();
			String line = br.readLine();

			while (line != null) {
				line = removeUTF8BOM(line);
				sb.append(line);
				sb.append(System.lineSeparator());
				line = br.readLine();
			}
			excludedGoodsListArea.setText(sb.toString());
		} catch (Exception e) {
			log.error(e.getLocalizedMessage(), e);
		}

		try {
			props.loadFromXML(new FileInputStream("calc_options.xml"));
		} catch (Exception e) {
			log.error(e.getLocalizedMessage(), e);
		}

		try {
			String val = props.getProperty(dayMoneyMinProp);
			if (val != null)
				spinnerMoneyFrom.setValue(Integer.parseInt(val));
		} catch (Exception e) {
			log.error(e.getLocalizedMessage(), e);
		}

		try {
			String val = props.getProperty(dayMoneyMaxProp);
			if (val != null)
				spinnerMoneyTo.setValue(Integer.parseInt(val));
		} catch (Exception e) {
			log.error(e.getLocalizedMessage(), e);
		}

		try {
			String val = props.getProperty(spinnerMaxGoodsQuantityProp);
			if (val != null)
				spinnerMaxGoodsQuantity.setValue(Integer.parseInt(val));
		} catch (Exception e) {
			log.error(e.getLocalizedMessage(), e);
		}

		try {
			String val = props.getProperty(checkBoxOpenFilesProp);
			if (val != null)
				checkBoxOpenFiles.setSelected(Boolean.parseBoolean(val));
		} catch (Exception e) {
			log.error(e.getLocalizedMessage(), e);
		}

		try {
			String val = props.getProperty(checkBoxDoNotChangeWarehouseProp);
			if (val != null)
				checkBoxDoNotChangeWarehouse.setSelected(Boolean.parseBoolean(val));
		} catch (Exception e) {
			log.error(e.getLocalizedMessage(), e);
		}
	}

	private void saveCalculationOptions() {
		String text = excludedGoodsListArea.getText();
		try (PrintWriter out = new PrintWriter(new File("excluded.txt"), "UTF-8")) {
			out.write(text);
			out.flush();
		} catch (Exception e) {
			log.error(e.getLocalizedMessage(), e);
		}

		try {
			Integer dayMoneyMin = (Integer) spinnerMoneyFrom.getValue();
			Integer dayMoneyMax = (Integer) spinnerMoneyTo.getValue();
			int quantity = (Integer) spinnerMaxGoodsQuantity.getValue();
			props.setProperty(dayMoneyMinProp, String.valueOf(dayMoneyMin));
			props.setProperty(dayMoneyMaxProp, String.valueOf(dayMoneyMax));
			props.setProperty(spinnerMaxGoodsQuantityProp, String.valueOf(quantity));
			props.setProperty(checkBoxOpenFilesProp, String.valueOf(checkBoxOpenFiles.isSelected()));
			props.setProperty(checkBoxDoNotChangeWarehouseProp,
					String.valueOf(checkBoxDoNotChangeWarehouse.isSelected()));
			props.storeToXML(new FileOutputStream("calc_options.xml"), "Calc options saved " + new Date());
		} catch (Exception e) {
			log.error(e.getLocalizedMessage(), e);
		}
	}

	private void enableControls(boolean enable) {
		warehouseFilePathField.setEnabled(enable);
		statementFilePathField.setEnabled(enable);
		btnCalculate.setEnabled(enable);
		spinnerMoneyTo.setEnabled(enable);
		spinnerMoneyFrom.setEnabled(enable);
		comboBoxMonth.setEnabled(enable);
		radioButtonFirst.setEnabled(enable);
		radioButtonFull.setEnabled(enable);
		radioButtonSecond.setEnabled(enable);
		spinnerYear.setEnabled(enable);
		checkBoxOpenFiles.setEnabled(enable);
		checkBoxDoNotChangeWarehouse.setEnabled(enable);
		btnSelectWarehouseFile.setEnabled(enable);
		btnSelectStatementFile.setEnabled(enable);
		excludedGoodsListArea.setEnabled(enable);
		spinnerMaxGoodsQuantity.setEnabled(enable);
	}

	private String selectSaveFile() {
		JFileChooser fc = new JFileChooser(prevDir);
		fc.setFileSelectionMode(JFileChooser.FILES_ONLY);
		fc.setFileFilter(new FileFilter() {

			@Override
			public String getDescription() {
				return "Microsoft Excel";
			}

			@Override
			public boolean accept(File f) {
				if (f.isDirectory()) {
					return true;
				}
				if (f.getName().trim().toLowerCase().endsWith(".xls")) {
					return true;
				}
				return false;
			}
		});
		int res = fc.showSaveDialog(frame);
		if (res == JFileChooser.APPROVE_OPTION) {

			File file = fc.getSelectedFile();
			String path = file.getAbsolutePath();
			if (!path.toLowerCase().endsWith(".xls")) {
				path = path + ".xls";
			}
			return path;
		}
		return null;
	}

	private File[] selectFolder(boolean multiSelectAllow) {
		JFileChooser fc = new JFileChooser(prevDir);
		fc.setFileSelectionMode(JFileChooser.FILES_ONLY);
		fc.setFileFilter(new FileFilter() {

			@Override
			public String getDescription() {
				return "Microsoft Excel";
			}

			@Override
			public boolean accept(File f) {
				if (f.isDirectory()) {
					return true;
				}
				if (f.getName().trim().toLowerCase().endsWith(".xls")) {
					return true;
				}
				if (f.getName().trim().toLowerCase().endsWith(".xlsx")) {
					return true;
				}
				return false;
			}
		});
		fc.setMultiSelectionEnabled(multiSelectAllow);
		int res = fc.showOpenDialog(frame);
		if (res == JFileChooser.APPROVE_OPTION) {
			if (multiSelectAllow) {
				File[] files = fc.getSelectedFiles();
				prevDir = files[0].getParentFile().getAbsolutePath();
				return files;
			} else {

				File file = fc.getSelectedFile();
				prevDir = file.getParentFile().getAbsolutePath();
				File[] files = new File[1];
				files[0] = file;
				return files;
			}
		}
		return null;
	}

	private static Image getIcon() {
		if (iconImage == null) {
			try {
				iconImage = ImageIO.read(MainWindow.class.getResource(iconName));
			} catch (IOException e) {
				log.error(e.getLocalizedMessage(), e);
				e.printStackTrace();
			}
		}
		return iconImage;
	}

	private File validatePath(String path, String fileType) {
		File file;
		if (path != null && path.trim().length() > 0) {
			file = new File(path);
			if (!file.exists()) {
				JOptionPane.showMessageDialog(frame, "Файл не найден: " + path, "Ошибка", JOptionPane.ERROR_MESSAGE);
				return null;
			}
			if (file.isDirectory()) {
				JOptionPane.showMessageDialog(frame, "Необходимо выбрать файл, а не папку: " + path, "Ошибка",
						JOptionPane.ERROR_MESSAGE);
				return null;
			}

			if (!file.canWrite()) {
				JOptionPane.showMessageDialog(frame, "Файл заблокирован на запись: " + path, "Ошибка",
						JOptionPane.ERROR_MESSAGE);
				return null;
			}

		} else {
			JOptionPane.showMessageDialog(frame, "Файл " + fileType + " не выбран!", "Ошибка",
					JOptionPane.ERROR_MESSAGE);
			return null;
		}
		return file;
	}

	private void calculate(boolean isFirstPart, boolean isOpenExcel) {
		// validate input

		// validate selected folders:
		File warehouseFile = validatePath(warehouseFilePathField.getText(), "остатков");
		File statementsFile = validatePath(statementFilePathField.getText(), "ведомости");
		if (warehouseFile == null || statementsFile == null) {
			return;
		}

		// validate periods

		Integer dayMoneyMin = (Integer) spinnerMoneyFrom.getValue();
		Integer dayMoneyMax = (Integer) spinnerMoneyTo.getValue();
		if (dayMoneyMin >= dayMoneyMax) {
			JOptionPane.showMessageDialog(frame, "Не правильный диапазон выручки!", "Ошибка",
					JOptionPane.ERROR_MESSAGE);
			return;
		}

		ArrayList<String> excludedGoodsList = new ArrayList<>();

		StringTokenizer st = new StringTokenizer(excludedGoodsListArea.getText().trim(), ";");
		while (st.hasMoreTokens()) {
			String exclusion = st.nextToken();
			if (exclusion.trim().length() > 0) {
				excludedGoodsList.add(exclusion.trim());
			}
		}

		SimpleDateFormat df = new SimpleDateFormat("MMMMM yyyyг", new Locale("ru"));
		int selectedMonth = comboBoxMonth.getSelectedIndex();
		Date yearValue = (Date) spinnerYear.getValue();
		Calendar calendar = Calendar.getInstance(new Locale("ru"));
		calendar.setTime(yearValue);
		calendar.set(Calendar.MONTH, selectedMonth);

		saveCalculationOptions();
		// log.debug(calendar.getTime());
		// read warehouse file
		InputStream myxlsWarehause = null;
		try {
			myxlsWarehause = new FileInputStream(warehouseFile);
		} catch (FileNotFoundException e) {
			log.error(e.getLocalizedMessage(), e);
			JOptionPane.showMessageDialog(frame, "Файл не найден: " + warehouseFile.getAbsolutePath(), "Ошибка",
					JOptionPane.ERROR_MESSAGE);
			return;
		}
		Workbook goodsWorkBook = null;
		try {
			goodsWorkBook = WorkbookFactory.create(myxlsWarehause);
		} catch (Exception e) {
			log.error(e.getLocalizedMessage(), e);
			JOptionPane.showMessageDialog(frame,
					"Невозможно открыть файл остатков - " + warehouseFile
							+ "\\nВозможно он поврежден.\nОткройте его в MS Excel и пересохраните!" + "\nОшибка: "
							+ e.getLocalizedMessage(),
					"Ошибка", JOptionPane.ERROR_MESSAGE);
			return;
		}
		TreeSet<Goods> goodsList = readGoods(excludedGoodsList, goodsWorkBook);
		if (goodsList == null) {
			return;
		}
		for (Goods goods : goodsList) {
			log.debug(goods.getName() + " quantity: " + goods.getTotalQuantity());
			TreeSet<GoodsPrice> prices = goods.getPriceList();
			for (GoodsPrice goodsPrice : prices) {
				log.debug("Row=" + goodsPrice.getWharehouseRowIndex() + " quantity=" + goodsPrice.getQuantity()
						+ " price=" + goodsPrice.getSellPrice());
			}
			log.debug("-----------");
		}

		// read statements file
		InputStream myxlsStatements = null;
		try {
			myxlsStatements = new FileInputStream(statementsFile);
		} catch (FileNotFoundException e) {
			log.error(e.getLocalizedMessage(), e);
			JOptionPane.showMessageDialog(frame, "Файл не найден: " + warehouseFile.getAbsolutePath(), "Ошибка",
					JOptionPane.ERROR_MESSAGE);
			return;
		}
		Workbook statementsWorkBook = null;
		try {
			statementsWorkBook = WorkbookFactory.create(myxlsStatements);
		} catch (Exception e1) {
			log.error(e1.getLocalizedMessage(), e1);
			JOptionPane.showMessageDialog(frame,
					"Невозможно открыть файл ведомости - " + statementsFile
							+ "\nВозможно он поврежден.\nОткройте его в MS Excel и пересохраните!" + "\nОшибка: "
							+ e1.getLocalizedMessage(),
					"Ошибка", JOptionPane.ERROR_MESSAGE);
			return;
		}

		Sheet sheet1 = statementsWorkBook.getSheetAt(0);
		Sheet sheet2 = statementsWorkBook.getSheetAt(1);
		Sheet sheetTotal = statementsWorkBook.getSheetAt(2);
		TreeSet<Goods> selectedGoods = null;
		if (isFirstPart) {
			// clean quantity for calculated dates:
			clearSheetValues(sheet1);
			clearSheetValues(sheet2);

			clearGoodsList(sheetTotal);

			selectedGoods = getGoodsForNewMonth(dayMoneyMin, dayMoneyMax, goodsList, calendar);
			if (selectedGoods == null) {
				return;
			}
			// save goods list into total sheet:
			int i = 0;
			ArrayList<Goods> selectedGoodsIndexed = new ArrayList<Goods>();
			for (Goods goods : selectedGoods) {
				selectedGoodsIndexed.add(goods);
				Row row = sheetTotal.getRow(i + 4);
				goods.setStatementRowIndex(i + 4);
				Cell nameCell = row.getCell(0);
				Cell sellPriceCell = row.getCell(1);
				Cell buyPriceCell = row.getCell(6);
				nameCell.setCellValue(goods.getName());
				sellPriceCell.setCellValue(goods.getPriceList().last().getSellPrice());
				buyPriceCell.setCellValue(goods.getPriceList().last().getBuyPriceWaVAT());
				i++;
			}

			boolean result = fillGoods(dayMoneyMin, dayMoneyMax, sheet1, selectedGoodsIndexed, 16, false, calendar);
			if (!result) {
				return;
			}
		} else {
			clearSheetValues(sheet2);

			try {
				selectedGoods = getGoodsForSecondMonthPart(goodsList, sheetTotal);
			} catch (Exception e) {
				log.error(e.getLocalizedMessage(), e);
				JOptionPane.showMessageDialog(frame,
						"Невозможно прочитать файл ведомости! Ошибка при чтении списка используемых товаров:\n"
								+ e.getClass().getSimpleName() + " - " + e.getLocalizedMessage() + "\n"
								+ getStackTrace(e),
						"Ошибка", JOptionPane.ERROR_MESSAGE);
				return;
			}
			if (selectedGoods == null) {
				return;
			}
			if (selectedGoods.size() == 0) {
				JOptionPane.showMessageDialog(frame,
						"Непредвиденная ошибка! На складе нет товаров\nкоторые использовались в первой половине месяца!",
						"Ошибка", JOptionPane.ERROR_MESSAGE);
				return;
			}
			// verify that total price of selected goods is enough for our month
			double totalPrice = 0;
			for (Goods goods : selectedGoods) {
				totalPrice += goods.getPriceList().last().getSellPrice() * goods.getPriceList().last().getQuantity();
			}
			log.debug("Added " + selectedGoods.size() + " goods with highest quantity, total goods price="
					+ Goods.round(totalPrice));
			// log.debug(calendar.getTime());
			int maxDayToFill = calendar.getActualMaximum(Calendar.DAY_OF_MONTH);
			int daysToFillQuantity = 0;
			for (int k = 1; k < maxDayToFill - 14; k++) {
				int day = k + 15;
				// exclude Sunday
				calendar.set(Calendar.DAY_OF_MONTH, day);
				// log.debug(calendar.getTime());
				if (calendar.get(Calendar.DAY_OF_WEEK) != Calendar.SUNDAY) {
					daysToFillQuantity++;
				}
			}
			log.debug("daysToFillQuantity=" + daysToFillQuantity);
			double requiredMoney = ((dayMoneyMax + dayMoneyMin) / 2.0) * daysToFillQuantity;
			if (totalPrice < requiredMoney) {
				JOptionPane.showMessageDialog(frame,
						"Непредвиденная ошибка! На складе не достаточное количество товаров\n"
								+ "которые использовались в первой половине месяца!\n" + "На складе товаров на сумму "
								+ Goods.round(totalPrice) + " грн.\nНеобходимо товаров на сумму "
								+ Goods.round(requiredMoney) + " грн",
						"Ошибка", JOptionPane.ERROR_MESSAGE);
				return;
			}

			ArrayList<Goods> selectedGoodsIndexed = new ArrayList<Goods>(selectedGoods);
			// log.debug(calendar.getTime());
			boolean result = fillGoods(dayMoneyMin, dayMoneyMax, sheet2, selectedGoodsIndexed, maxDayToFill - 14, true,
					calendar);
			// log.debug(calendar.getTime());
			if (!result) {
				return;
			}
		}

		// hide rows with no goods:
		if (isFirstPart) {
			log.debug("Hide rows without goods");
			for (int j = 4; j < sheet1.getLastRowNum(); j++) {
				Row row = sheetTotal.getRow(j);
				Row row1 = sheet1.getRow(j);
				Row row2 = sheet2.getRow(j);
				Cell c = row1.getCell(0);
				try {
					if (c != null) {
						String val = getCellValue(c);
						if (val != null) {
							if (val.equalsIgnoreCase("Итого")) {
								break;
							}
						}
					}
				} catch (Exception e) {
					log.error("Error hidding rows: " + e.getLocalizedMessage(), e);
				}
				if (j - 4 < selectedGoods.size()) {
					row.setZeroHeight(false);
					row1.setZeroHeight(false);
					row2.setZeroHeight(false);
				} else {
					row.setZeroHeight(true);
					row1.setZeroHeight(true);
					row2.setZeroHeight(true);
				}
			}
		}

		// update date
		sheetTotal.getRow(2).getCell(2).setCellValue(df.format(calendar.getTime()));
		log.debug("Update date");

		if (statementsWorkBook instanceof HSSFWorkbook) {
			HSSFFormulaEvaluator.evaluateAllFormulaCells((HSSFWorkbook) statementsWorkBook);
			log.debug("evaluateAllFormulaCells");
		} else if (statementsWorkBook instanceof XSSFWorkbook) {
			XSSFFormulaEvaluator.evaluateAllFormulaCells((XSSFWorkbook) statementsWorkBook);
			log.debug("evaluateAllFormulaCells");
		}

		try {
			// save workbook
			log.debug("save workbook");
			FileOutputStream fileOut = new FileOutputStream(statementsFile);
			statementsWorkBook.write(fileOut);
			fileOut.close();
			log.debug("save done");
			if (isOpenExcel) {
				try {
					Desktop desktop = null;
					if (Desktop.isDesktopSupported()) {
						desktop = Desktop.getDesktop();
					}
					if (desktop != null) {
						desktop.open(statementsFile);
					}
				} catch (IOException ioe) {
					log.error("Cannot open report in default viewer", ioe);
				}
			}

		} catch (Exception e) {
			log.error(e.getLocalizedMessage(), e);
			JOptionPane
					.showMessageDialog(frame,
							"Ошибка при сохранении файла ведомости!\n" + e.getClass().getSimpleName() + " - "
									+ e.getLocalizedMessage() + "\n" + getStackTrace(e),
							"Ошибка", JOptionPane.ERROR_MESSAGE);
		}

		if (!checkBoxDoNotChangeWarehouse.isSelected()) {
			if (selectedGoods != null) {
				Sheet sheet = goodsWorkBook.getSheetAt(0);
				for (Goods goods : selectedGoods) {
					TreeSet<GoodsPrice> priceList = goods.getPriceList();
					for (GoodsPrice goodsPrice : priceList) {
						Row row = sheet.getRow(goodsPrice.getWharehouseRowIndex());
						Cell cell = row.getCell(3);
						cell.setCellValue(goodsPrice.getQuantity());
						cell = row.getCell(4);
						cell.setCellValue(goodsPrice.getPriceWVAT());
					}
				}
				try {
					FileOutputStream fileOut = new FileOutputStream(warehouseFile);
					goodsWorkBook.write(fileOut);
					fileOut.close();
				} catch (Exception e) {
					log.error(e.getLocalizedMessage(), e);
					JOptionPane.showMessageDialog(frame,
							"Ошибка при сохранении файла остатков!\n" + e.getClass().getSimpleName() + " - "
									+ e.getLocalizedMessage() + "\n" + getStackTrace(e),
							"Ошибка", JOptionPane.ERROR_MESSAGE);
				}
			}
		}
	}

	private boolean fillGoods(Integer dayMoneyMin, Integer dayMoneyMax, Sheet sheet,
			ArrayList<Goods> selectedGoodsIndexed, int maxDayToFill, boolean secondPart, Calendar calendar) {
		// fill with data:
		int goodsQuantity = selectedGoodsIndexed.size();
		Random r = new Random();
		ArrayList<Integer> notFilledDays = new ArrayList<Integer>();
		for (int k = 1; k < maxDayToFill; k++) {
			double totalPrice = 0;
			int day = k;
			if (secondPart) {
				day += 15;
			}
			// exclude Sunday
			calendar.set(Calendar.DAY_OF_MONTH, day);
			if (calendar.get(Calendar.DAY_OF_WEEK) == Calendar.SUNDAY) {
				log.debug("Skip Sunday " + calendar.getTime());
				continue;
			}
			int randomDaySum = dayMoneyMin + (int) (Math.random() * (((dayMoneyMax - 5) - dayMoneyMin) + 1));
			int whileCounter = 0;
			while (totalPrice < randomDaySum && whileCounter < 50000) {
				int randomGoodsIndex = r.nextInt(goodsQuantity);
				whileCounter++;

				Goods randomGoods = selectedGoodsIndexed.get(randomGoodsIndex);
				// log.debug("Random good index: " + randomGoodsIndex);
				if (totalPrice + randomGoods.getPriceList().last().getSellPrice() < dayMoneyMax
						&& randomGoods.getTotalQuantity() > 0) {
					Row row = sheet.getRow(randomGoods.getStatementRowIndex());
					Cell c = row.getCell(getColumnIndexByMonthDay(k));
					try {
						String currentVal = getCellValue(c);
						Integer currentIntVal = new Integer(1);
						if (currentVal != null && currentVal.trim().length() > 0) {
							currentIntVal = (int) Math.round(Double.parseDouble(currentVal));
							if (currentIntVal >= 3) {
								continue;
							}
							currentIntVal++;
						}
						totalPrice += randomGoods.getPriceList().last().getSellPrice();
						c.setCellValue(currentIntVal);
						randomGoods.decreaseQuantity();
						log.debug("Day " + day + ". add \"" + randomGoods.getName() + "\" price="
								+ randomGoods.getPriceList().last().getSellPrice() + " quantity=" + currentIntVal
								+ " totalPrice=" + Goods.round(totalPrice) + " cell="
								+ new CellReference(c).formatAsString());
					} catch (Exception e) {
						log.error(e.getLocalizedMessage(), e);
						JOptionPane.showMessageDialog(frame,
								"Ошибка при работе с файлом ведомости!\n" + e.getClass().getSimpleName() + " - "
										+ e.getLocalizedMessage() + "\n" + getStackTrace(e),
								"Ошибка", JOptionPane.ERROR_MESSAGE);
						return false;

					}
				} else {
					log.debug("Day " + day + ". skip \"" + randomGoods.getName() + "\" price="
							+ randomGoods.getPriceList().last().getSellPrice() + " totalPrice="
							+ Goods.round(totalPrice));
				}
			}
			if (whileCounter > 45000) {
				notFilledDays.add(day);

			}
			log.debug("Filled day " + day + ". Got price=" + Goods.round(totalPrice) + " needed: " + randomDaySum);
			log.debug("------------------------------------------------------------");
		}
		if (!notFilledDays.isEmpty()) {
			JOptionPane.showMessageDialog(frame,
					"Не дотаточно товаров для " + notFilledDays + " дней!\nПроверьте файл ведомости", "Внимание",
					JOptionPane.WARNING_MESSAGE);
		}
		return true;
	}

	private TreeSet<Goods> getGoodsForSecondMonthPart(TreeSet<Goods> goodsList, Sheet sheetTotal) throws Exception {
		// get used goods from Excel
		TreeSet<Goods> selectedGoods = new TreeSet<>();

		for (int j = 4; j < totalStatementRow; j++) {
			Row row = sheetTotal.getRow(j);
			Cell c = row.getCell(0);
			if (c != null) {
				String name = getCellValue(c);
				if (name != null && name.trim().length() > 0) {
					// find this goods in our list
					for (Goods goods : goodsList) {
						if (goods.getName().equals(name)) {
							goods.setStatementRowIndex(j);
							selectedGoods.add(goods);
						}
					}
				}
			}
		}

		return selectedGoods;
	}

	private TreeSet<Goods> getGoodsForNewMonth(Integer dayMoneyMin, Integer dayMoneyMax, TreeSet<Goods> goodsList,
			Calendar calendar) {
		// get 30 goods with highest quantity
		ArrayList<String> selectedNames = new ArrayList<>();
		TreeSet<Goods> selectedGoods = new TreeSet<>();
		int quantity = (Integer) spinnerMaxGoodsQuantity.getValue();
		for (Goods g : goodsList) {
			if (selectedGoods.size() < quantity) {
				if (!selectedNames.contains(g.getName())) {
					selectedGoods.add(g);
					selectedNames.add(g.getName());
				}
			} else {
				break;
			}
		}

		// verify that total price of selected goods is enough for our month
		double totalPrice = 0;
		for (Goods goods : selectedGoods) {
			totalPrice += goods.getPriceList().last().getSellPrice() * goods.getTotalQuantity();
		}
		log.debug("Added " + quantity + " goods with highest quantity, total goods price=" + Goods.round(totalPrice));

		int daysToFillQuantity = 0;
		for (int k = 1; k < 16; k++) {
			// exclude Sunday
			calendar.set(Calendar.DAY_OF_MONTH, k);
			if (calendar.get(Calendar.DAY_OF_WEEK) != Calendar.SUNDAY) {
				daysToFillQuantity++;
			}
		}
		log.debug("daysToFillQuantity=" + daysToFillQuantity);
		double requiredMoney = ((dayMoneyMax + dayMoneyMin) / 2.0) * daysToFillQuantity;
		log.debug("We need goods for total price: " + Goods.round(requiredMoney));
		if (totalPrice < requiredMoney) {
			// add more goods
			for (Goods g : goodsList) {
				if (totalPrice < requiredMoney) {
					if (!selectedGoods.contains(g) && selectedGoods.size() < 75) {
						selectedGoods.add(g);
						totalPrice += g.getPriceList().last().getSellPrice() * g.getTotalQuantity();
					}
				} else {
					break;
				}
			}
			log.debug("Added additional goods to have required total price. Quantity: " + selectedGoods.size()
					+ " Total goods price: " + Goods.round(totalPrice));
		}
		if (totalPrice < requiredMoney) {
			JOptionPane.showMessageDialog(frame,
					"Непредвиденная ошибка! На складе не достаточное количество товаров!\n"
							+ "На складе товаров на сумму " + Goods.round(totalPrice)
							+ " грн.\nНеобходимо товаров на сумму " + Goods.round(requiredMoney) + " грн",
					"Ошибка", JOptionPane.ERROR_MESSAGE);
			return null;
		}
		return selectedGoods;
	}

	private void clearGoodsList(Sheet sheetTotal) {
		for (int j = 4; j < totalStatementRow; j++) {
			// clear goods name
			Row row = sheetTotal.getRow(j);
			Cell c = row.getCell(0);
			if (c != null)
				c.setCellType(Cell.CELL_TYPE_BLANK);
			// clear sell price
			c = row.getCell(1);
			if (c != null)
				c.setCellType(Cell.CELL_TYPE_BLANK);
			// clear buy price
			c = row.getCell(6);
			if (c != null)
				c.setCellType(Cell.CELL_TYPE_BLANK);

		}
	}

	private void clearSheetValues(Sheet sheet) {
		for (int i = 1; i < 17; i++) {
			for (int j = 4; j < totalStatementRow; j++) {
				Row row = sheet.getRow(j);
				Cell c = row.getCell(getColumnIndexByMonthDay(i));
				if (c != null)
					c.setCellType(Cell.CELL_TYPE_BLANK);
			}
		}
	}

	private TreeSet<Goods> readGoods(ArrayList<String> excludedGoodsList, Workbook wb) {
		try {
			TreeSet<Goods> goodsList = new TreeSet<>();
			Sheet sheet = wb.getSheetAt(0);
			boolean foundHeader = false;
			String currentCategory = null;
			for (Row row : sheet) {
				if (!foundHeader) {
					Cell cell = row.getCell(0);
					String val = getCellValue(cell);
					if (val != null) {
						if (val.equals("Склад/ТМЦ/Партия")) {
							foundHeader = true;
						}
					}
				} else {
					// read goods:
					Cell cell = row.getCell(0);
					String name = getCellValue(cell);
					String cellName = "";
					if (name != null && name.trim().length() > 0) {
						if (!excludedGoodsList.contains(name)) {
							cell = row.getCell(3);
							if (cell != null) {
								cellName = " (ячейка " + new CellReference(cell).formatAsString() + ")";
							} else {
								cellName = "";
							}
							String quantityStr = getCellValue(cell);
							if (quantityStr != null && quantityStr.trim().length() > 0) {
								try {
									Double quantity = Double.parseDouble(quantityStr);
									Cell cell2 = row.getCell(4);
									if (cell2 != null) {
										cellName = " (ячейка " + new CellReference(cell2).formatAsString() + ")";
									} else {
										cellName = "";
									}
									String priceStr = getCellValue(cell2);
									if (priceStr != null && priceStr.trim().length() > 0) {
										try {
											Double price = Double.parseDouble(priceStr);
											Goods g = new Goods(name, currentCategory);
											boolean found = false;
											for (Goods gEl : goodsList) {
												if (gEl.equals(g)) {
													gEl.addPriceElement(Math.round(quantity.floatValue()),
															price.doubleValue(), row.getRowNum());
													found = true;
													break;
												}
											}
											if (!found) {
												g.addPriceElement(Math.round(quantity.floatValue()),
														price.doubleValue(), row.getRowNum());
												goodsList.add(g);
											}
										} catch (NumberFormatException e) {
											log.error(e.getLocalizedMessage(), e);
											JOptionPane.showMessageDialog(frame,
													"Ошибка чтения файла остатков!\nЦена товара \"" + name
															+ "\" не цифра:" + priceStr + "\nСтрока "
															+ (row.getRowNum() + 1) + cellName,
													"Ошибка", JOptionPane.ERROR_MESSAGE);
											return null;
										}
									} else {
										JOptionPane.showMessageDialog(frame,
												"Ошибка чтения файла остатков!\nЦена товара \"" + name + "\" не задано!"
														+ "\nСтрока " + (row.getRowNum() + 1) + cellName,
												"Ошибка", JOptionPane.ERROR_MESSAGE);
										return null;
									}
								} catch (NumberFormatException e) {
									log.error(e.getLocalizedMessage(), e);
									JOptionPane.showMessageDialog(frame,
											"Ошибка чтения файла остатков!\nКоличество товара \"" + name
													+ "\" не цифра:" + quantityStr + "\nСтрока " + (row.getRowNum() + 1)
													+ cellName,
											"Ошибка", JOptionPane.ERROR_MESSAGE);
									return null;
								}
							} else {
								JOptionPane.showMessageDialog(frame,
										"Ошибка чтения файла остатков!\nКоличество товара \"" + name + "\" не задано!"
												+ "\nСтрока " + (row.getRowNum() + 1) + cellName,
										"Ошибка", JOptionPane.ERROR_MESSAGE);
								return null;
							}
						} else {
							currentCategory = name;
						}

					} else {
						JOptionPane.showMessageDialog(frame,
								"Невозможно прочитать Excel файл остатков!\nНазвание товара в строке " + row.getRowNum()
										+ " не задано!",
								"Ошибка", JOptionPane.ERROR_MESSAGE);
						return null;
					}
				}
			}

			if (!foundHeader) {
				JOptionPane.showMessageDialog(frame,
						"Невозможно прочитать Excel файл остатков!\nСтолбец \"Склад/ТМЦ/Партия\" не найден!", "Ошибка",
						JOptionPane.ERROR_MESSAGE);
				return null;
			}
			return goodsList;
		} catch (Exception e) {
			log.error(e.getLocalizedMessage(), e);
			JOptionPane
					.showMessageDialog(frame,
							"Невозможно прочитать Excel файл остатков!\nОшибка: " + e.getClass().getSimpleName() + " - "
									+ e.getLocalizedMessage() + "\n" + getStackTrace(e),
							"Ошибка", JOptionPane.ERROR_MESSAGE);
			return null;
		}
	}

	private static String getStackTrace(Throwable aThrowable) {
		final Writer result = new StringWriter();
		final PrintWriter printWriter = new PrintWriter(result);
		aThrowable.printStackTrace(printWriter);
		return result.toString();
	}

	private String getCellValue(Cell cell) throws Exception {
		try {
			String result = null;
			if (cell != null) {
				// log.info("Read "+cell.getSheet().getSheetName()+":"+new
				// CellReference(cell).formatAsString());
				SimpleDateFormat dateTimeFormat = new SimpleDateFormat("dd.MM.yyyy HH:mm:ss");
				switch (cell.getCellType()) {
				case Cell.CELL_TYPE_BLANK:
					result = null;
					break;
				case Cell.CELL_TYPE_STRING:
					result = (cell.getRichStringCellValue().getString()).trim();
					break;
				case Cell.CELL_TYPE_NUMERIC:
					if (DateUtil.isCellDateFormatted(cell)) {
						result = dateTimeFormat.format(cell.getDateCellValue());
					} else {
						result = Double.toString(cell.getNumericCellValue());
					}
					break;
				case Cell.CELL_TYPE_BOOLEAN:
					result = Boolean.toString(cell.getBooleanCellValue());
					break;
				case Cell.CELL_TYPE_FORMULA:
					switch (cell.getCachedFormulaResultType()) {
					case Cell.CELL_TYPE_BOOLEAN:
						result = Boolean.toString(cell.getBooleanCellValue());
						break;
					case Cell.CELL_TYPE_NUMERIC:
						result = Double.toString(cell.getNumericCellValue());
						break;
					case Cell.CELL_TYPE_STRING:
						result = (cell.getStringCellValue()).trim();
						break;
					case Cell.CELL_TYPE_BLANK:
						break;
					case Cell.CELL_TYPE_ERROR:
						result = Byte.toString(cell.getErrorCellValue());
						break;

					// CELL_TYPE_FORMULA will never occur
					case Cell.CELL_TYPE_FORMULA:
						break;
					}

					break;
				default:
				}
			}
			return result;
		} catch (Exception e) {
			if (e.getMessage().indexOf("Could not resolve external workbook name") != -1
					|| e.getMessage().indexOf("not implemented yet") != -1) {
				// this is special POI exceptions because some cell has link to
				// other Workbook, but this functionality is not supported
				// generate own error message
				throw new Exception("Cannot read Excel file!\nProgram need to read value from cell "
						+ new CellReference(cell).formatAsString() + " sheet '" + cell.getSheet().getSheetName()
						+ "', but it refecenes to other external Excel file", e);
			} else {
				throw e;
			}
		}
	}

	private static int getColumnIndexByMonthDay(int day) {
		if (day < 7) {
			return day * 3 - 1;
		} else if (day < 13) {
			return day * 3 + 1;
		} else {
			return day * 3 + 3;
		}
	}

	// 1. Вихідні дні (неділя)
	// 2. Брать суммарну кількість однакового товару. На складі віднімать самий
	// перший. А в набивачці самій останній товар

}
