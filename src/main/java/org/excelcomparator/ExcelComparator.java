package org.excelcomparator;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;

/**
 * Utility to compare Excel File Contents cell by cell for all sheets.
 */
public class ExcelComparator {

	private static final String BOLD = "BOLD";
	private static final String BOTTOM_BORDER = "BOTTOM BORDER";
	private static final String BRACKET_END = "]";
	private static final String BRACKET_START = " [";
	private static final String CELL_ALIGNMENT_DOES_NOT_MATCH = "Cell Alignment does not Match ::";
	private static final String CELL_BORDER_ATTRIBUTES_DOES_NOT_MATCH = "Cell Border Attributes does not Match ::";
	private static final String CELL_DATA_DOES_NOT_MATCH = "Cell Data does not Match ::";
	private static final String CELL_DATA_TYPE_DOES_NOT_MATCH = "Cell Data-Type does not Match in :: ";
	private static final String CELL_FILL_COLOR_DOES_NOT_MATCH = "Cell Fill Color does not Match ::";
	private static final String CELL_FILL_PATTERN_DOES_NOT_MATCH = "Cell Fill pattern does not Match ::";
	private static final String CELL_FONT_ATTRIBUTES_DOES_NOT_MATCH = "Cell Font Attributes does not Match ::";
	private static final String CELL_FONT_FAMILY_DOES_NOT_MATCH = "Cell Font Family does not Match ::";
	private static final String CELL_FONT_SIZE_DOES_NOT_MATCH = "Cell Font Size does not Match ::";
	private static final String CELL_PROTECTION_DOES_NOT_MATCH = "Cell Protection does not Match ::";
	private static final String ITALICS = "ITALICS";
	private static final String LEFT_BORDER = "LEFT BORDER";
	private static final String LINE_SEPARATOR = "line.separator";
	private static final String NAME_OF_THE_SHEETS_DO_NOT_MATCH = "Name of the sheets do not match :: ";
	private static final String NEXT_STR = " -> ";
	private static final String NO_BOTTOM_BORDER = "NO BOTTOM BORDER";
	private static final String NO_COLOR = "NO COLOR";
	private static final String NO_LEFT_BORDER = "NO LEFT BORDER";
	private static final String NO_RIGHT_BORDER = "NO RIGHT BORDER";
	private static final String NO_TOP_BORDER = "NO TOP BORDER";
	private static final String NOT_BOLD = "NOT BOLD";
	private static final String NOT_EQUALS = " != ";
	private static final String NOT_ITALICS = "NOT ITALICS";
	private static final String NOT_UNDERLINE = "NOT UNDERLINE";
	private static final String NUMBER_OF_COLUMNS_DOES_NOT_MATCH = "Number Of Columns does not Match :: ";
	private static final String NUMBER_OF_ROWS_DOES_NOT_MATCH = "Number Of Rows does not Match :: ";
	private static final String NUMBER_OF_SHEETS_DO_NOT_MATCH = "Number of Sheets do not match :: ";
	private static final String RIGHT_BORDER = "RIGHT BORDER";
	private static final String TOP_BORDER = "TOP BORDER";
	private static final String UNDERLINE = "UNDERLINE";
	private static final String WORKBOOK1 = "workbook1";
	private static final String WORKBOOK2 = "workbook2";

	/**
	 * Utility to compare Excel File Contents cell by cell for all sheets. This
	 * method returns an object with a successful flag and list of differences.
	 *
	 * @param file1
	 *            the file1
	 * @param file2
	 *            the file2
	 * @return the excel file difference
	 */
	public static ExcelFileDifference compare(File file1, File file2) {

		ExcelFileDifference excelFileDifference;
		try {
			final Workbook workbook1 = WorkbookFactory.create(file1);
			final Workbook workbook2 = WorkbookFactory.create(file2);
			excelFileDifference = ExcelComparator.compare(workbook1, workbook2);
		} catch (final EncryptedDocumentException e) {
			excelFileDifference = new ExcelFileDifference();
			excelFileDifference.error = e;
		} catch (final InvalidFormatException e) {
			excelFileDifference = new ExcelFileDifference();
			excelFileDifference.error = e;
		} catch (final IOException e) {
			excelFileDifference = new ExcelFileDifference();
			excelFileDifference.error = e;
		}
		return excelFileDifference;

	}

	/**
	 * Utility to compare Excel File Contents cell by cell for all sheets. This
	 * method returns an object with a successful flag and list of differences.
	 *
	 * @param workbook1
	 *            the workbook1
	 * @param workbook2
	 *            the workbook2
	 * @return the Excel file difference containing a flag and a list of
	 *         differences
	 *
	 */
	public static ExcelFileDifference compare(Workbook workbook1, Workbook workbook2) {
		final List<String> listOfDifferences = ExcelComparator.compareWorkBookContents(workbook1, workbook2);
		return ExcelComparator.populateListOfDifferences(listOfDifferences);
	}

	/**
	 * Utility to compare Excel File Contents cell by cell for all sheets. This
	 * method only returns a success flag.
	 *
	 * @param file1
	 *            the file1
	 * @param file2
	 *            the file2
	 * @return true, if successful
	 */
	public static boolean compareWithoutDetails(File file1, File file2) {
		try {
			final Workbook workbook1 = WorkbookFactory.create(file1);
			final Workbook workbook2 = WorkbookFactory.create(file2);
			final ExcelFileDifference excelFileDifferences = ExcelComparator.compare(workbook1, workbook2);
			if ((excelFileDifferences != null) && excelFileDifferences.isDifferenceFound) {
				return false;
			}
		} catch (final FileNotFoundException e) {
			return false;
		} catch (final EncryptedDocumentException e) {
			return false;
		} catch (final InvalidFormatException e) {
			return false;
		} catch (final IOException e) {
			return false;
		}
		return true;
	}

	/**
	 * Compare work book contents.
	 *
	 * @param workbook1
	 *            the workbook1
	 * @param workbook2
	 *            the workbook2
	 * @return the list
	 */
	private static List<String> compareWorkBookContents(Workbook workbook1, Workbook workbook2) {
		final ExcelComparator excelComparator = new ExcelComparator();
		final List<String> listOfDifferences = new ArrayList<String>();
		excelComparator.compareNumberOfSheets(workbook1, workbook2, listOfDifferences);
		excelComparator.compareSheetNames(workbook1, workbook2, listOfDifferences);
		excelComparator.compareSheetData(workbook1, workbook2, listOfDifferences);
		return listOfDifferences;
	}

	/**
	 * Populate list of differences.
	 *
	 * @param listOfDifferences
	 *            the list of differences
	 * @return the excel file difference
	 */
	private static ExcelFileDifference populateListOfDifferences(List<String> listOfDifferences) {
		final ExcelFileDifference excelFileDifference = new ExcelFileDifference();
		excelFileDifference.isDifferenceFound = listOfDifferences.size() > 0;
		excelFileDifference.listOfDifferences = listOfDifferences;
		return excelFileDifference;
	}

	/**
	 * Compare data in all sheets.
	 *
	 * @param workbook1
	 *            the workbook1
	 * @param workbook2
	 *            the workbook2
	 * @param listOfDifferences
	 *            the list of differences
	 */
	private void compareDataInAllSheets(Workbook workbook1, Workbook workbook2, List<String> listOfDifferences) {
		for (int i = 0; i < workbook1.getNumberOfSheets(); i++) {
			final Sheet sheetWorkBook1 = workbook1.getSheetAt(i);
			Sheet sheetWorkBook2;
			if (workbook2.getNumberOfSheets() > i) {
				sheetWorkBook2 = workbook2.getSheetAt(i);
			} else {
				sheetWorkBook2 = null;
			}

			for (int j = 0; j < sheetWorkBook1.getPhysicalNumberOfRows(); j++) {
				final Row rowWorkBook1 = sheetWorkBook1.getRow(j);
				Row rowWorkBook2;
				if (sheetWorkBook2 != null) {
					rowWorkBook2 = sheetWorkBook2.getRow(j);
				} else {
					rowWorkBook2 = null;
				}

				if ((rowWorkBook1 == null) || (rowWorkBook2 == null)) {
					continue;
				}
				for (int k = 0; k < rowWorkBook1.getLastCellNum(); k++) {
					final Cell cellWorkBook1 = rowWorkBook1.getCell(k);
					final Cell cellWorkBook2 = rowWorkBook2.getCell(k);

					if (!((null == cellWorkBook1) || (null == cellWorkBook2))) {
						if (isCellTypeMatches(cellWorkBook1, cellWorkBook2)) {

							listOfDifferences.add(getMessage(workbook1, workbook2, i, cellWorkBook1, cellWorkBook2,
									ExcelComparator.CELL_DATA_TYPE_DOES_NOT_MATCH, cellWorkBook1.getCellType() + "",
									cellWorkBook2.getCellType() + ""));
						}

						if (isCellContentTypeBlank(cellWorkBook1)) {
							if (isCellContentMatches(cellWorkBook1, cellWorkBook2)) {

								listOfDifferences.add(getMessage(workbook1, workbook2, i, cellWorkBook1, cellWorkBook2,
										ExcelComparator.CELL_DATA_DOES_NOT_MATCH,
										cellWorkBook1.getRichStringCellValue() + "",
										cellWorkBook2.getRichStringCellValue() + ""));

							}

						} else if (isCellContentTypeBoolean(cellWorkBook1)) {
							if (isCellContentMatchesForBoolean(cellWorkBook1, cellWorkBook2)) {
								listOfDifferences.add(getMessage(workbook1, workbook2, i, cellWorkBook1, cellWorkBook2,
										ExcelComparator.CELL_DATA_DOES_NOT_MATCH,
										cellWorkBook1.getBooleanCellValue() + "",
										cellWorkBook2.getBooleanCellValue() + ""));

							}

						} else if (isCellContentInError(cellWorkBook1)) {
							if (isCellContentMatches(cellWorkBook1, cellWorkBook2)) {

								listOfDifferences.add(getMessage(workbook1, workbook2, i, cellWorkBook1, cellWorkBook2,
										ExcelComparator.CELL_DATA_DOES_NOT_MATCH,
										cellWorkBook1.getRichStringCellValue() + "",
										cellWorkBook2.getRichStringCellValue() + ""));

							}
						} else if (isCellContentFormula(cellWorkBook1)) {
							if (isCellContentMatchesForFormula(cellWorkBook1, cellWorkBook2)) {

								listOfDifferences.add(getMessage(workbook1, workbook2, i, cellWorkBook1, cellWorkBook2,
										ExcelComparator.CELL_DATA_DOES_NOT_MATCH, cellWorkBook1.getCellFormula() + "",
										cellWorkBook2.getCellFormula() + ""));

							}

						} else if (isCellContentTypeNumeric(cellWorkBook1)) {
							if (DateUtil.isCellDateFormatted(cellWorkBook1)) {
								if (isCellContentMatchesForDate(cellWorkBook1, cellWorkBook2)) {
									listOfDifferences.add(getMessage(workbook1, workbook2, i, cellWorkBook1,
											cellWorkBook2, ExcelComparator.CELL_DATA_DOES_NOT_MATCH,
											cellWorkBook1.getDateCellValue() + "",
											cellWorkBook2.getDateCellValue() + ""));

								}
							} else {
								if (isCellContentMatchesForNumeric(cellWorkBook1, cellWorkBook2)) {
									listOfDifferences.add(getMessage(workbook1, workbook2, i, cellWorkBook1,
											cellWorkBook2, ExcelComparator.CELL_DATA_DOES_NOT_MATCH,
											cellWorkBook1.getNumericCellValue() + "",
											cellWorkBook2.getNumericCellValue() + ""));

								}
							}

						} else if (isCellContentTypeString(cellWorkBook1)) {
							if (isCellContentMatches(cellWorkBook1, cellWorkBook2)) {
								listOfDifferences.add(getMessage(workbook1, workbook2, i, cellWorkBook1, cellWorkBook2,
										ExcelComparator.CELL_DATA_DOES_NOT_MATCH,
										cellWorkBook1.getRichStringCellValue().getString(),
										cellWorkBook2.getRichStringCellValue().getString()));
							}
						}

						if (isCellFillPatternMatches(cellWorkBook1, cellWorkBook2)) {
							listOfDifferences.add(getMessage(workbook1, workbook2, i, cellWorkBook1, cellWorkBook2,
									ExcelComparator.CELL_FILL_PATTERN_DOES_NOT_MATCH,
									cellWorkBook1.getCellStyle().getFillPattern() + "",
									cellWorkBook2.getCellStyle().getFillPattern() + ""));

						}

						if (isCellAlignmentMatches(cellWorkBook1, cellWorkBook2)) {
							listOfDifferences.add(getMessage(workbook1, workbook2, i, cellWorkBook1, cellWorkBook2,
									ExcelComparator.CELL_ALIGNMENT_DOES_NOT_MATCH,
									cellWorkBook1.getRichStringCellValue().getString(),
									cellWorkBook2.getRichStringCellValue().getString()));

						}

						if (isCellHiddenMatches(cellWorkBook1, cellWorkBook2)) {
							listOfDifferences.add(getMessage(workbook1, workbook2, i, cellWorkBook1, cellWorkBook2,
									ExcelComparator.CELL_PROTECTION_DOES_NOT_MATCH,
									cellWorkBook1.getCellStyle().getHidden() ? "HIDDEN" : "NOT HIDDEN",
									cellWorkBook2.getCellStyle().getHidden() ? "HIDDEN" : "NOT HIDDEN"));

						}

						if (isCellLockedMatches(cellWorkBook1, cellWorkBook2)) {
							listOfDifferences.add(getMessage(workbook1, workbook2, i, cellWorkBook1, cellWorkBook2,
									ExcelComparator.CELL_PROTECTION_DOES_NOT_MATCH,
									cellWorkBook1.getCellStyle().getLocked() ? "LOCKED" : "NOT LOCKED",
									cellWorkBook2.getCellStyle().getLocked() ? "LOCKED" : "NOT LOCKED"));

						}

						if (isCellFontFamilyMatches(cellWorkBook1, cellWorkBook2)) {
							listOfDifferences.add(getMessage(workbook1, workbook2, i, cellWorkBook1, cellWorkBook2,
									ExcelComparator.CELL_FONT_FAMILY_DOES_NOT_MATCH,
									((XSSFCellStyle) cellWorkBook1.getCellStyle()).getFont().getFontName(),
									((XSSFCellStyle) cellWorkBook2.getCellStyle()).getFont().getFontName()));

						}

						if (isCellFontSizeMatches(cellWorkBook1, cellWorkBook2)) {
							listOfDifferences.add(getMessage(workbook1, workbook2, i, cellWorkBook1, cellWorkBook2,
									ExcelComparator.CELL_FONT_SIZE_DOES_NOT_MATCH,
									((XSSFCellStyle) cellWorkBook1.getCellStyle()).getFont().getFontHeightInPoints()
											+ "",
									((XSSFCellStyle) cellWorkBook2.getCellStyle()).getFont().getFontHeightInPoints()
											+ ""));

						}

						if (isCellFontBoldMatches(cellWorkBook1, cellWorkBook2)) {
							listOfDifferences.add(getMessage(workbook1, workbook2, i, cellWorkBook1, cellWorkBook2,
									ExcelComparator.CELL_FONT_ATTRIBUTES_DOES_NOT_MATCH,
									((XSSFCellStyle) cellWorkBook1.getCellStyle()).getFont().getBold()
											? ExcelComparator.BOLD : ExcelComparator.NOT_BOLD,
									((XSSFCellStyle) cellWorkBook2.getCellStyle()).getFont().getBold()
											? ExcelComparator.BOLD : ExcelComparator.NOT_BOLD));

						}

						if (isCellUnderLineMatches(cellWorkBook1, cellWorkBook2)) {
							listOfDifferences.add(getMessage(workbook1, workbook2, i, cellWorkBook1, cellWorkBook2,
									ExcelComparator.CELL_FONT_ATTRIBUTES_DOES_NOT_MATCH,
									((XSSFCellStyle) cellWorkBook1.getCellStyle()).getFont().getUnderline() == 1
											? ExcelComparator.UNDERLINE : ExcelComparator.NOT_UNDERLINE,
									((XSSFCellStyle) cellWorkBook2.getCellStyle()).getFont().getUnderline() == 1
											? ExcelComparator.UNDERLINE : ExcelComparator.NOT_UNDERLINE));

						}

						if (isCellFontItalicsMatches(cellWorkBook1, cellWorkBook2)) {
							listOfDifferences.add(getMessage(workbook1, workbook2, i, cellWorkBook1, cellWorkBook2,
									ExcelComparator.CELL_FONT_ATTRIBUTES_DOES_NOT_MATCH,
									((XSSFCellStyle) cellWorkBook1.getCellStyle()).getFont().getItalic()
											? ExcelComparator.ITALICS : ExcelComparator.NOT_ITALICS,
									((XSSFCellStyle) cellWorkBook2.getCellStyle()).getFont().getItalic()
											? ExcelComparator.ITALICS : ExcelComparator.NOT_ITALICS));

						}

						if (isCellBorderBottomMatches(cellWorkBook1, cellWorkBook2)) {
							listOfDifferences.add(getMessage(workbook1, workbook2, i, cellWorkBook1, cellWorkBook2,
									ExcelComparator.CELL_BORDER_ATTRIBUTES_DOES_NOT_MATCH,
									((XSSFCellStyle) cellWorkBook1.getCellStyle()).getBorderBottom() == 1
											? ExcelComparator.BOTTOM_BORDER : ExcelComparator.NO_BOTTOM_BORDER,
									((XSSFCellStyle) cellWorkBook2.getCellStyle()).getBorderBottom() == 1
											? ExcelComparator.BOTTOM_BORDER : ExcelComparator.NO_BOTTOM_BORDER));

						}

						if (isCellBorderLeftMatches(cellWorkBook1, cellWorkBook2)) {
							listOfDifferences.add(getMessage(workbook1, workbook2, i, cellWorkBook1, cellWorkBook2,
									ExcelComparator.CELL_BORDER_ATTRIBUTES_DOES_NOT_MATCH,
									((XSSFCellStyle) cellWorkBook1.getCellStyle()).getBorderLeft() == 1
											? ExcelComparator.LEFT_BORDER : ExcelComparator.NO_LEFT_BORDER,
									((XSSFCellStyle) cellWorkBook2.getCellStyle()).getBorderLeft() == 1
											? ExcelComparator.LEFT_BORDER : ExcelComparator.NO_LEFT_BORDER));

						}

						if (isCellBorderRightMatches(cellWorkBook1, cellWorkBook2)) {
							listOfDifferences.add(getMessage(workbook1, workbook2, i, cellWorkBook1, cellWorkBook2,
									ExcelComparator.CELL_BORDER_ATTRIBUTES_DOES_NOT_MATCH,
									((XSSFCellStyle) cellWorkBook1.getCellStyle()).getBorderRight() == 1
											? ExcelComparator.RIGHT_BORDER : ExcelComparator.NO_RIGHT_BORDER,
									((XSSFCellStyle) cellWorkBook2.getCellStyle()).getBorderRight() == 1
											? ExcelComparator.RIGHT_BORDER : ExcelComparator.NO_RIGHT_BORDER));

						}

						if (isCellBorderTopMatches(cellWorkBook1, cellWorkBook2)) {
							listOfDifferences.add(getMessage(workbook1, workbook2, i, cellWorkBook1, cellWorkBook2,
									ExcelComparator.CELL_BORDER_ATTRIBUTES_DOES_NOT_MATCH,
									((XSSFCellStyle) cellWorkBook1.getCellStyle()).getBorderTop() == 1
											? ExcelComparator.TOP_BORDER : ExcelComparator.NO_TOP_BORDER,
									((XSSFCellStyle) cellWorkBook2.getCellStyle()).getBorderTop() == 1
											? ExcelComparator.TOP_BORDER : ExcelComparator.NO_TOP_BORDER));

						}

						if (isCellBackGroundFillMatchesAndEmpty(cellWorkBook1, cellWorkBook2)) {
							continue;
						} else if (isCellFillBackGroundMatchesAndEitherEmpty(cellWorkBook1, cellWorkBook2)) {
							listOfDifferences.add(getMessage(workbook1, workbook2, i, cellWorkBook1, cellWorkBook2,
									ExcelComparator.CELL_FILL_COLOR_DOES_NOT_MATCH, ExcelComparator.NO_COLOR,
									((XSSFColor) cellWorkBook2.getCellStyle().getFillForegroundColorColor())
											.getARGBHex() + ""));

						} else if (isCellFillBackGroundMatchesAndSecondEmpty(cellWorkBook1, cellWorkBook2)) {
							listOfDifferences.add(getMessage(workbook1, workbook2, i, cellWorkBook1, cellWorkBook2,
									ExcelComparator.CELL_FILL_COLOR_DOES_NOT_MATCH,
									((XSSFColor) cellWorkBook1.getCellStyle().getFillForegroundColorColor())
											.getARGBHex() + "",
									ExcelComparator.NO_COLOR));

						} else {
							if (isCellFileBackGroundMatches(cellWorkBook1, cellWorkBook2)) {
								listOfDifferences.add(getMessage(workbook1, workbook2, i, cellWorkBook1, cellWorkBook2,
										ExcelComparator.CELL_FILL_COLOR_DOES_NOT_MATCH,
										((XSSFColor) cellWorkBook1.getCellStyle().getFillForegroundColorColor())
												.getARGBHex() + "",
										((XSSFColor) cellWorkBook2.getCellStyle().getFillForegroundColorColor())
												.getARGBHex() + ""));

							}
						}

					}
				}
			}
		}
	}

	/**
	 * Compare number of columns in sheets.
	 *
	 * @param workbook1
	 *            the workbook1
	 * @param workbook2
	 *            the workbook2
	 * @param listOfDifferences
	 *            the list of differences
	 */
	private void compareNumberOfColumnsInSheets(Workbook workbook1, Workbook workbook2,
			List<String> listOfDifferences) {
		for (int i = 0; i < workbook1.getNumberOfSheets(); i++) {
			final Sheet sheetWorkBook1 = workbook1.getSheetAt(i);
			Sheet sheetWorkBook2;
			if (workbook2.getNumberOfSheets() > i) {
				sheetWorkBook2 = workbook2.getSheetAt(i);
			} else {
				sheetWorkBook2 = null;
			}
			if (isWorkBookEmpty(sheetWorkBook1, sheetWorkBook2)) {
				if (isNumberOfColumnsMatches(sheetWorkBook1, sheetWorkBook2)) {
					String noOfCols;
					String sheetName;
					if (sheetWorkBook2 != null) {
						noOfCols = sheetWorkBook2.getRow(0).getLastCellNum() + "";
						sheetName = workbook2.getSheetName(i);
					} else {
						noOfCols = "";
						sheetName = "";
					}
					final short lastCellNumForWbk1 = sheetWorkBook1.getRow(0) != null
							? sheetWorkBook1.getRow(0).getLastCellNum() : 0;
					listOfDifferences.add(ExcelComparator.NUMBER_OF_COLUMNS_DOES_NOT_MATCH
							+ System.getProperty(ExcelComparator.LINE_SEPARATOR) + ExcelComparator.WORKBOOK1
							+ ExcelComparator.NEXT_STR + workbook1.getSheetName(i) + ExcelComparator.NEXT_STR
							+ lastCellNumForWbk1 + ExcelComparator.NOT_EQUALS + ExcelComparator.WORKBOOK2
							+ ExcelComparator.NEXT_STR + sheetName + ExcelComparator.NEXT_STR + noOfCols);
				}
			}
		}
	}

	/**
	 * Compare number of rows in sheets.
	 *
	 * @param workbook1
	 *            the workbook1
	 * @param workbook2
	 *            the workbook2
	 * @param listOfDifferences
	 *            the list of differences
	 */
	private void compareNumberOfRowsInSheets(Workbook workbook1, Workbook workbook2, List<String> listOfDifferences) {
		for (int i = 0; i < workbook1.getNumberOfSheets(); i++) {
			final Sheet sheetWorkBook1 = workbook1.getSheetAt(i);
			Sheet sheetWorkBook2;
			if (workbook2.getNumberOfSheets() > i) {
				sheetWorkBook2 = workbook2.getSheetAt(i);
			} else {
				sheetWorkBook2 = null;
			}
			if (isNumberOfRowsMatches(sheetWorkBook1, sheetWorkBook2)) {
				String noOfRows;
				String sheetName;
				if (sheetWorkBook2 != null) {
					noOfRows = sheetWorkBook2.getPhysicalNumberOfRows() + "";
					sheetName = workbook2.getSheetName(i);
				} else {
					noOfRows = "";
					sheetName = "";
				}
				listOfDifferences.add(ExcelComparator.NUMBER_OF_ROWS_DOES_NOT_MATCH
						+ System.getProperty(ExcelComparator.LINE_SEPARATOR) + ExcelComparator.WORKBOOK1
						+ ExcelComparator.NEXT_STR + workbook1.getSheetName(i) + ExcelComparator.NEXT_STR
						+ sheetWorkBook1.getPhysicalNumberOfRows() + ExcelComparator.NOT_EQUALS
						+ ExcelComparator.WORKBOOK2 + ExcelComparator.NEXT_STR + sheetName + ExcelComparator.NEXT_STR
						+ noOfRows);
			}
		}

	}

	/**
	 * Compare number of sheets.
	 *
	 * @param workbook1
	 *            the workbook1
	 * @param workbook2
	 *            the workbook2
	 * @param listOfDifferences
	 *            the list of differences
	 */
	private void compareNumberOfSheets(Workbook workbook1, Workbook workbook2, List<String> listOfDifferences) {
		if (isNumberOfSheetsMatches(workbook1, workbook2)) {
			listOfDifferences.add(ExcelComparator.NUMBER_OF_SHEETS_DO_NOT_MATCH
					+ System.getProperty(ExcelComparator.LINE_SEPARATOR) + ExcelComparator.WORKBOOK1
					+ ExcelComparator.NEXT_STR + workbook1.getNumberOfSheets() + ExcelComparator.NOT_EQUALS
					+ ExcelComparator.WORKBOOK2 + ExcelComparator.NEXT_STR + workbook2.getNumberOfSheets());
		}
	}

	/**
	 * Compare sheet data.
	 *
	 * @param workbook1
	 *            the workbook1
	 * @param workbook2
	 *            the workbook2
	 * @param listOfDifferences
	 *            the list of differences
	 */
	private void compareSheetData(Workbook workbook1, Workbook workbook2, List<String> listOfDifferences) {
		compareNumberOfRowsInSheets(workbook1, workbook2, listOfDifferences);
		compareNumberOfColumnsInSheets(workbook1, workbook2, listOfDifferences);
		compareDataInAllSheets(workbook1, workbook2, listOfDifferences);

	}

	/**
	 * Compare sheet names.
	 *
	 * @param workbook1
	 *            the workbook1
	 * @param workbook2
	 *            the workbook2
	 * @param listOfDifferences
	 *            the list of differences
	 */
	private void compareSheetNames(Workbook workbook1, Workbook workbook2, List<String> listOfDifferences) {
		for (int i = 0; i < workbook1.getNumberOfSheets(); i++) {
			if (isNameOfSheetMatches(workbook1, workbook2, i)) {
				final String sheetname = workbook2.getNumberOfSheets() > i ? workbook2.getSheetName(i) : "";
				listOfDifferences.add(ExcelComparator.NAME_OF_THE_SHEETS_DO_NOT_MATCH
						+ System.getProperty(ExcelComparator.LINE_SEPARATOR) + ExcelComparator.WORKBOOK1
						+ ExcelComparator.NEXT_STR + workbook1.getSheetName(i) + ExcelComparator.BRACKET_START + (i + 1)
						+ ExcelComparator.BRACKET_END + ExcelComparator.NOT_EQUALS + ExcelComparator.WORKBOOK2
						+ ExcelComparator.NEXT_STR + sheetname + ExcelComparator.BRACKET_START + (i + 1)
						+ ExcelComparator.BRACKET_END);
			}
		}
	}

	/**
	 * Gets the message.
	 *
	 * @param workbook1
	 *            the workbook1
	 * @param workbook2
	 *            the workbook2
	 * @param i
	 *            the i
	 * @param cellWorkBook1
	 *            the cell work book1
	 * @param cellWorkBook2
	 *            the cell work book2
	 * @param messageStart
	 *            the message start
	 * @param workBook1Value
	 *            the work book1 value
	 * @param workBook2Value
	 *            the work book2 value
	 * @return the message
	 */
	private String getMessage(Workbook workbook1, Workbook workbook2, int i, Cell cellWorkBook1, Cell cellWorkBook2,
			String messageStart, String workBook1Value, String workBook2Value) {
		final StringBuilder sb = new StringBuilder();
		return sb.append(messageStart).append(System.getProperty(ExcelComparator.LINE_SEPARATOR))
				.append(ExcelComparator.WORKBOOK1).append(ExcelComparator.NEXT_STR).append(workbook1.getSheetName(i))
				.append(ExcelComparator.NEXT_STR)
				.append(new CellReference(cellWorkBook1.getRowIndex(), cellWorkBook1.getColumnIndex()).formatAsString())
				.append(ExcelComparator.BRACKET_START).append(workBook1Value).append(ExcelComparator.BRACKET_END)
				.append(ExcelComparator.NOT_EQUALS).append(ExcelComparator.WORKBOOK2).append(ExcelComparator.NEXT_STR)
				.append(workbook2.getSheetName(i)).append(ExcelComparator.NEXT_STR)
				.append(new CellReference(cellWorkBook2.getRowIndex(), cellWorkBook2.getColumnIndex()).formatAsString())
				.append(ExcelComparator.BRACKET_START).append(workBook2Value).append(ExcelComparator.BRACKET_END)
				.toString();
	}

	/**
	 * Checks if cell alignment matches.
	 *
	 * @param cellWorkBook1
	 *            the cell work book1
	 * @param cellWorkBook2
	 *            the cell work book2
	 * @return true, if cell alignment matches
	 */
	private boolean isCellAlignmentMatches(Cell cellWorkBook1, Cell cellWorkBook2) {
		return cellWorkBook1.getCellStyle().getAlignment() != cellWorkBook2.getCellStyle().getAlignment();
	}

	/**
	 * Checks if cell back ground fill matches and empty.
	 *
	 * @param cellWorkBook1
	 *            the cell work book1
	 * @param cellWorkBook2
	 *            the cell work book2
	 * @return true, if cell back ground fill matches and empty
	 */
	private boolean isCellBackGroundFillMatchesAndEmpty(Cell cellWorkBook1, Cell cellWorkBook2) {
		return (cellWorkBook1.getCellStyle().getFillForegroundColorColor() == null)
				&& (cellWorkBook2.getCellStyle().getFillForegroundColorColor() == null);
	}

	/**
	 * Checks if cell border bottom matches.
	 *
	 * @param cellWorkBook1
	 *            the cell work book1
	 * @param cellWorkBook2
	 *            the cell work book2
	 * @return true, if cell border bottom matches
	 */
	private boolean isCellBorderBottomMatches(Cell cellWorkBook1, Cell cellWorkBook2) {
		if (cellWorkBook1.getCellStyle() instanceof XSSFCellStyle) {
			return ((XSSFCellStyle) cellWorkBook1.getCellStyle())
					.getBorderBottom() != ((XSSFCellStyle) cellWorkBook2.getCellStyle()).getBorderBottom();
		} else {
			return false;
		}

	}

	/**
	 * Checks if cell border left matches.
	 *
	 * @param cellWorkBook1
	 *            the cell work book1
	 * @param cellWorkBook2
	 *            the cell work book2
	 * @return true, if cell border left matches
	 */
	private boolean isCellBorderLeftMatches(Cell cellWorkBook1, Cell cellWorkBook2) {
		if (cellWorkBook1.getCellStyle() instanceof XSSFCellStyle) {
			return ((XSSFCellStyle) cellWorkBook1.getCellStyle())
					.getBorderLeft() != ((XSSFCellStyle) cellWorkBook2.getCellStyle()).getBorderLeft();
		} else {
			return false;
		}

	}

	/**
	 * Checks if cell border right matches.
	 *
	 * @param cellWorkBook1
	 *            the cell work book1
	 * @param cellWorkBook2
	 *            the cell work book2
	 * @return true, if cell border right matches
	 */
	private boolean isCellBorderRightMatches(Cell cellWorkBook1, Cell cellWorkBook2) {
		if (cellWorkBook1.getCellStyle() instanceof XSSFCellStyle) {
			return ((XSSFCellStyle) cellWorkBook1.getCellStyle())
					.getBorderRight() != ((XSSFCellStyle) cellWorkBook2.getCellStyle()).getBorderRight();
		} else {
			return false;
		}

	}

	/**
	 * Checks if cell border top matches.
	 *
	 * @param cellWorkBook1
	 *            the cell work book1
	 * @param cellWorkBook2
	 *            the cell work book2
	 * @return true, if cell border top matches
	 */
	private boolean isCellBorderTopMatches(Cell cellWorkBook1, Cell cellWorkBook2) {
		if (cellWorkBook1.getCellStyle() instanceof XSSFCellStyle) {
			return ((XSSFCellStyle) cellWorkBook1.getCellStyle())
					.getBorderTop() != ((XSSFCellStyle) cellWorkBook2.getCellStyle()).getBorderTop();
		} else {
			return false;
		}

	}

	/**
	 * Checks if cell content formula.
	 *
	 * @param cellWorkBook1
	 *            the cell work book1
	 * @return true, if cell content formula
	 */
	private boolean isCellContentFormula(Cell cellWorkBook1) {
		return cellWorkBook1.getCellType() == Cell.CELL_TYPE_FORMULA;
	}

	/**
	 * Checks if cell content in error.
	 *
	 * @param cellWorkBook1
	 *            the cell work book1
	 * @return true, if cell content in error
	 */
	private boolean isCellContentInError(Cell cellWorkBook1) {
		return cellWorkBook1.getCellType() == Cell.CELL_TYPE_ERROR;
	}

	/**
	 * Checks if cell content matches.
	 *
	 * @param cellWorkBook1
	 *            the cell work book1
	 * @param cellWorkBook2
	 *            the cell work book2
	 * @return true, if cell content matches
	 */
	private boolean isCellContentMatches(Cell cellWorkBook1, Cell cellWorkBook2) {
		return !(cellWorkBook1.getRichStringCellValue().getString()
				.equals(cellWorkBook2.getRichStringCellValue().getString()));
	}

	/**
	 * Checks if cell content matches for boolean.
	 *
	 * @param cellWorkBook1
	 *            the cell work book1
	 * @param cellWorkBook2
	 *            the cell work book2
	 * @return true, if cell content matches for boolean
	 */
	private boolean isCellContentMatchesForBoolean(Cell cellWorkBook1, Cell cellWorkBook2) {
		return !(cellWorkBook1.getBooleanCellValue() == cellWorkBook2.getBooleanCellValue());
	}

	/**
	 * Checks if cell content matches for date.
	 *
	 * @param cellWorkBook1
	 *            the cell work book1
	 * @param cellWorkBook2
	 *            the cell work book2
	 * @return true, if cell content matches for date
	 */
	private boolean isCellContentMatchesForDate(Cell cellWorkBook1, Cell cellWorkBook2) {
		return !(cellWorkBook1.getDateCellValue().equals(cellWorkBook2.getDateCellValue()));
	}

	/**
	 * Checks if cell content matches for formula.
	 *
	 * @param cellWorkBook1
	 *            the cell work book1
	 * @param cellWorkBook2
	 *            the cell work book2
	 * @return true, if cell content matches for formula
	 */
	private boolean isCellContentMatchesForFormula(Cell cellWorkBook1, Cell cellWorkBook2) {
		return !(cellWorkBook1.getCellFormula().equals(cellWorkBook2.getCellFormula()));
	}

	/**
	 * Checks if cell content matches for numeric.
	 *
	 * @param cellWorkBook1
	 *            the cell work book1
	 * @param cellWorkBook2
	 *            the cell work book2
	 * @return true, if cell content matches for numeric
	 */
	private boolean isCellContentMatchesForNumeric(Cell cellWorkBook1, Cell cellWorkBook2) {
		return !(cellWorkBook1.getNumericCellValue() == cellWorkBook2.getNumericCellValue());
	}

	/**
	 * Checks if cell content type blank.
	 *
	 * @param cellWorkBook1
	 *            the cell work book1
	 * @return true, if cell content type blank
	 */
	private boolean isCellContentTypeBlank(Cell cellWorkBook1) {
		return cellWorkBook1.getCellType() == Cell.CELL_TYPE_BLANK;
	}

	/**
	 * Checks if cell content type boolean.
	 *
	 * @param cellWorkBook1
	 *            the cell work book1
	 * @return true, if cell content type boolean
	 */
	private boolean isCellContentTypeBoolean(Cell cellWorkBook1) {
		return cellWorkBook1.getCellType() == Cell.CELL_TYPE_BOOLEAN;
	}

	/**
	 * Checks if cell content type numeric.
	 *
	 * @param cellWorkBook1
	 *            the cell work book1
	 * @return true, if cell content type numeric
	 */
	private boolean isCellContentTypeNumeric(Cell cellWorkBook1) {
		return cellWorkBook1.getCellType() == Cell.CELL_TYPE_NUMERIC;
	}

	/**
	 * Checks if cell content type string.
	 *
	 * @param cellWorkBook1
	 *            the cell work book1
	 * @return true, if cell content type string
	 */
	private boolean isCellContentTypeString(Cell cellWorkBook1) {
		return cellWorkBook1.getCellType() == Cell.CELL_TYPE_STRING;
	}

	/**
	 * Checks if cell file back ground matches.
	 *
	 * @param cellWorkBook1
	 *            the cell work book1
	 * @param cellWorkBook2
	 *            the cell work book2
	 * @return true, if cell file back ground matches
	 */
	private boolean isCellFileBackGroundMatches(Cell cellWorkBook1, Cell cellWorkBook2) {
		if (cellWorkBook1.getCellStyle() instanceof XSSFCellStyle) {
			return !((XSSFColor) cellWorkBook1.getCellStyle().getFillForegroundColorColor()).getARGBHex()
					.equals(((XSSFColor) cellWorkBook2.getCellStyle().getFillForegroundColorColor()).getARGBHex());
		} else {
			return false;
		}

	}

	/**
	 * Checks if cell fill back ground matches and either empty.
	 *
	 * @param cellWorkBook1
	 *            the cell work book1
	 * @param cellWorkBook2
	 *            the cell work book2
	 * @return true, if cell fill back ground matches and either empty
	 */
	private boolean isCellFillBackGroundMatchesAndEitherEmpty(Cell cellWorkBook1, Cell cellWorkBook2) {
		return (cellWorkBook1.getCellStyle().getFillForegroundColorColor() == null)
				&& (cellWorkBook2.getCellStyle().getFillForegroundColorColor() != null);
	}

	/**
	 * Checks if cell fill back ground matches and second empty.
	 *
	 * @param cellWorkBook1
	 *            the cell work book1
	 * @param cellWorkBook2
	 *            the cell work book2
	 * @return true, if cell fill back ground matches and second empty
	 */
	private boolean isCellFillBackGroundMatchesAndSecondEmpty(Cell cellWorkBook1, Cell cellWorkBook2) {
		return (cellWorkBook1.getCellStyle().getFillForegroundColorColor() != null)
				&& (cellWorkBook2.getCellStyle().getFillForegroundColorColor() == null);
	}

	/**
	 * Checks if cell fill pattern matches.
	 *
	 * @param cellWorkBook1
	 *            the cell work book1
	 * @param cellWorkBook2
	 *            the cell work book2
	 * @return true, if cell fill pattern matches
	 */
	private boolean isCellFillPatternMatches(Cell cellWorkBook1, Cell cellWorkBook2) {
		return cellWorkBook1.getCellStyle().getFillPattern() != cellWorkBook2.getCellStyle().getFillPattern();
	}

	/**
	 * Checks if cell font bold matches.
	 *
	 * @param cellWorkBook1
	 *            the cell work book1
	 * @param cellWorkBook2
	 *            the cell work book2
	 * @return true, if cell font bold matches
	 */
	private boolean isCellFontBoldMatches(Cell cellWorkBook1, Cell cellWorkBook2) {
		if (cellWorkBook1.getCellStyle() instanceof XSSFCellStyle) {
			return ((XSSFCellStyle) cellWorkBook1.getCellStyle()).getFont()
					.getBold() != ((XSSFCellStyle) cellWorkBook2.getCellStyle()).getFont().getBold();
		} else {
			return false;
		}

	}

	/**
	 * Checks if cell font family matches.
	 *
	 * @param cellWorkBook1
	 *            the cell work book1
	 * @param cellWorkBook2
	 *            the cell work book2
	 * @return true, if cell font family matches
	 */
	private boolean isCellFontFamilyMatches(Cell cellWorkBook1, Cell cellWorkBook2) {
		if (cellWorkBook1.getCellStyle() instanceof XSSFCellStyle) {
			return !(((XSSFCellStyle) cellWorkBook1.getCellStyle()).getFont().getFontName()
					.equals(((XSSFCellStyle) cellWorkBook2.getCellStyle()).getFont().getFontName()));
		} else {
			return false;
		}
	}

	/**
	 * Checks if cell font italics matches.
	 *
	 * @param cellWorkBook1
	 *            the cell work book1
	 * @param cellWorkBook2
	 *            the cell work book2
	 * @return true, if cell font italics matches
	 */
	private boolean isCellFontItalicsMatches(Cell cellWorkBook1, Cell cellWorkBook2) {
		if (cellWorkBook1.getCellStyle() instanceof XSSFCellStyle) {
			return ((XSSFCellStyle) cellWorkBook1.getCellStyle()).getFont()
					.getItalic() != ((XSSFCellStyle) cellWorkBook2.getCellStyle()).getFont().getItalic();
		} else {
			return false;
		}

	}

	/**
	 * Checks if cell font size matches.
	 *
	 * @param cellWorkBook1
	 *            the cell work book1
	 * @param cellWorkBook2
	 *            the cell work book2
	 * @return true, if cell font size matches
	 */
	private boolean isCellFontSizeMatches(Cell cellWorkBook1, Cell cellWorkBook2) {
		if (cellWorkBook1.getCellStyle() instanceof XSSFCellStyle) {
			return ((XSSFCellStyle) cellWorkBook1.getCellStyle()).getFont()
					.getFontHeightInPoints() != ((XSSFCellStyle) cellWorkBook2.getCellStyle()).getFont()
							.getFontHeightInPoints();
		} else {
			return false;
		}

	}

	/**
	 * Checks if cell hidden matches.
	 *
	 * @param cellWorkBook1
	 *            the cell work book1
	 * @param cellWorkBook2
	 *            the cell work book2
	 * @return true, if cell hidden matches
	 */
	private boolean isCellHiddenMatches(Cell cellWorkBook1, Cell cellWorkBook2) {
		return cellWorkBook1.getCellStyle().getHidden() != cellWorkBook2.getCellStyle().getHidden();
	}

	/**
	 * Checks if cell locked matches.
	 *
	 * @param cellWorkBook1
	 *            the cell work book1
	 * @param cellWorkBook2
	 *            the cell work book2
	 * @return true, if cell locked matches
	 */
	private boolean isCellLockedMatches(Cell cellWorkBook1, Cell cellWorkBook2) {
		return cellWorkBook1.getCellStyle().getLocked() != cellWorkBook2.getCellStyle().getLocked();
	}

	/**
	 * Checks if cell type matches.
	 *
	 * @param cellWorkBook1
	 *            the cell work book1
	 * @param cellWorkBook2
	 *            the cell work book2
	 * @return true, if cell type matches
	 */
	private boolean isCellTypeMatches(Cell cellWorkBook1, Cell cellWorkBook2) {
		return !(cellWorkBook1.getCellType() == cellWorkBook2.getCellType());
	}

	/**
	 * Checks if cell under line matches.
	 *
	 * @param cellWorkBook1
	 *            the cell work book1
	 * @param cellWorkBook2
	 *            the cell work book2
	 * @return true, if cell under line matches
	 */
	private boolean isCellUnderLineMatches(Cell cellWorkBook1, Cell cellWorkBook2) {
		if (cellWorkBook1.getCellStyle() instanceof XSSFCellStyle) {
			return ((XSSFCellStyle) cellWorkBook1.getCellStyle()).getFont()
					.getUnderline() != ((XSSFCellStyle) cellWorkBook2.getCellStyle()).getFont().getUnderline();
		} else {
			return false;
		}

	}

	/**
	 * Checks if name of sheet matches.
	 *
	 * @param workbook1
	 *            the workbook1
	 * @param workbook2
	 *            the workbook2
	 * @param i
	 *            the i
	 * @return true, if name of sheet matches
	 */
	private boolean isNameOfSheetMatches(Workbook workbook1, Workbook workbook2, int i) {
		if (workbook2.getNumberOfSheets() > i) {
			return !(workbook1.getSheetName(i).equals(workbook2.getSheetName(i)));
		} else {
			return true;
		}

	}

	/**
	 * Checks if number of columns matches.
	 *
	 * @param sheetWorkBook1
	 *            the sheet work book1
	 * @param sheetWorkBook2
	 *            the sheet work book2
	 * @return true, if number of columns matches
	 */
	private boolean isNumberOfColumnsMatches(Sheet sheetWorkBook1, Sheet sheetWorkBook2) {
		if (sheetWorkBook2 != null) {
			return !(sheetWorkBook1.getRow(0).getLastCellNum() == sheetWorkBook2.getRow(0).getLastCellNum());
		} else {
			return true;
		}

	}

	/**
	 * Checks if number of rows matches.
	 *
	 * @param sheetWorkBook1
	 *            the sheet work book1
	 * @param sheetWorkBook2
	 *            the sheet work book2
	 * @return true, if number of rows matches
	 */
	private boolean isNumberOfRowsMatches(Sheet sheetWorkBook1, Sheet sheetWorkBook2) {
		if (sheetWorkBook2 != null) {
			return !(sheetWorkBook1.getPhysicalNumberOfRows() == sheetWorkBook2.getPhysicalNumberOfRows());
		} else {
			return true;
		}

	}

	/**
	 * Checks if number of sheets matches.
	 *
	 * @param workbook1
	 *            the workbook1
	 * @param workbook2
	 *            the workbook2
	 * @return true, if number of sheets matches
	 */
	private boolean isNumberOfSheetsMatches(Workbook workbook1, Workbook workbook2) {
		return !(workbook1.getNumberOfSheets() == workbook2.getNumberOfSheets());
	}

	/**
	 * Checks if is work book empty.
	 *
	 * @param sheetWorkBook1
	 *            the sheet work book1
	 * @param sheetWorkBook2
	 *            the sheet work book2
	 * @return true, if is work book empty
	 */
	private boolean isWorkBookEmpty(Sheet sheetWorkBook1, Sheet sheetWorkBook2) {
		if (sheetWorkBook2 != null) {
			return !((null == sheetWorkBook1.getRow(0)) || (null == sheetWorkBook2.getRow(0)));
		} else {
			return true;
		}

	}

}

class ExcelFileDifference {
	Exception error;
	boolean isDifferenceFound;
	List<String> listOfDifferences;
}
