package com.apache.poi.excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFCreationHelper;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XLSReader {

	private String pathXLSFile;
	private XSSFWorkbook xssfWorkbook = null;

	/**
	 * Constructor XLS
	 * 
	 * @param pathXLSFile
	 * @throws FileNotFoundException
	 * @throws IOException
	 */
	public XLSReader(String pathXLSFile) throws FileNotFoundException, IOException {

		this.pathXLSFile = pathXLSFile;
		xssfWorkbook = new XSSFWorkbook(new FileInputStream(pathXLSFile));
		// xssfSheet = xssfWorkbook.getSheetAt(0);

	}

	/**
	 * if excel Sheet does not exist return -1 else return number rows in Excel
	 * Sheet
	 * 
	 * @param sheetName
	 * @return
	 */
	public int getRowCount(String sheetName) {

		int indexSheet = xssfWorkbook.getSheetIndex(sheetName);
		if (!isSheetExist(sheetName)) {
			return -1;
		}

		return xssfWorkbook.getSheetAt(indexSheet).getLastRowNum() == 0
				&& xssfWorkbook.getSheetAt(indexSheet).getRow(0) == null ? 0
						: xssfWorkbook.getSheetAt(indexSheet).getLastRowNum() + 1;
	}

	/**
	 * method Close File XLS
	 * 
	 * @throws IOException
	 */
	public void close() throws IOException {
		xssfWorkbook.close();
	}

	/**
	 * 
	 * @param sheetName
	 * @param colName
	 * @param rowNum
	 * @return
	 */
	public String getCellData(String sheetName, String colName, int rowNum) {

		if (rowNum <= 0) {
			return null;
		}

		if (!isSheetExist(sheetName)) {
			return null;
		}

		int index = xssfWorkbook.getSheetIndex(sheetName);
		XSSFSheet sheet = xssfWorkbook.getSheetAt(index);
		int colNum = -1;
		XSSFRow row = sheet.getRow(0);
		for (int i = 0; i < row.getLastCellNum(); i++) {
			// System.out.println(row.getCell(i).getStringCellValue());
			if (row.getCell(i).getStringCellValue().trim().equalsIgnoreCase(colName.trim())) {
				colNum = i;
				break;
			}
		}

		if (colNum == -1) {
			return null;
		}

		row = sheet.getRow(rowNum - 1);
		if (row == null) {
			return null;
		}

		XSSFCell cell = row.getCell(colNum);

		if (cell == null) {
			return "";
		}

		return getCellValue(cell);

	}

	/**
	 * 
	 * @param sheetName
	 * @param colNum
	 * @param rowNum
	 * @return
	 */
	public String getCellData(String sheetName, int colNum, int rowNum) {

		if (rowNum <= 0) {
			return null;
		}

		if (!isSheetExist(sheetName)) {
			return null;
		}

		int index = xssfWorkbook.getSheetIndex(sheetName);
		XSSFSheet sheet = xssfWorkbook.getSheetAt(index);
		XSSFRow row = sheet.getRow(rowNum - 1);
		if (row == null) {
			return null;
		}

		XSSFCell cell = row.getCell(colNum - 1);

		if (cell == null) {
			return "";
		}

		return getCellValue(cell);
	}

	/**
	 * 
	 * @param sheetName
	 * @return
	 */
	public boolean isSheetExist(String sheetName) {
		return xssfWorkbook.getSheetIndex(sheetName) == -1 ? false : true;
	}

	/**
	 * 
	 * @param sheetName
	 * @return
	 */
	public int getColumnCount(String sheetName) {
		if (!isSheetExist(sheetName)) {
			return -1;
		}

		int indexSheet = xssfWorkbook.getSheetIndex(sheetName);

		return xssfWorkbook.getSheetAt(indexSheet).getLastRowNum() == 0
				&& xssfWorkbook.getSheetAt(indexSheet).getRow(0) == null ? 0
						: xssfWorkbook.getSheetAt(indexSheet).getRow(0).getLastCellNum();
	}

	/**
	 * 
	 * @param sheetName
	 * @param colName
	 * @param cellValue
	 * @return
	 */
	public int getCellRowNum(String sheetName, String colName, String cellValue) {
		for (int i = 2; i <= getRowCount(sheetName); i++) {
			if (getCellData(sheetName, colName, i).equalsIgnoreCase(cellValue)) {
				return i;
			}
		}
		return -1;
	}

	/**
	 * returns true if sheet is created successfully else false
	 * 
	 * @param sheetname
	 * @return
	 * @throws IOException
	 */
	public boolean addSheet(String sheetname) {

		FileOutputStream fileOutputStream = null;

		xssfWorkbook.createSheet(sheetname);
		try {
			fileOutputStream = new FileOutputStream(pathXLSFile);
			xssfWorkbook.write(fileOutputStream);
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		} finally {
			try {
				if (fileOutputStream != null)
					fileOutputStream.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}

		return true;
	}

	/**
	 * returns true if sheet is removed successfully else false if sheet does not
	 * exist
	 * 
	 * @param sheetName
	 * @return
	 */
	public boolean removeSheet(String sheetName) {
		if (!isSheetExist(sheetName)) {
			return false;
		}

		int index = xssfWorkbook.getSheetIndex(sheetName);
		FileOutputStream fileOutputStream = null;
		try {
			xssfWorkbook.removeSheetAt(index);
			fileOutputStream = new FileOutputStream(pathXLSFile);
			xssfWorkbook.write(fileOutputStream);
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		} finally {

			try {
				if (fileOutputStream != null)
					fileOutputStream.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		return true;
	}

	/**
	 * returns true if column is created successfully
	 * 
	 * @param sheetName
	 * @param colName
	 * @return
	 */
	public boolean addColumn(String sheetName, String colName) {

		FileInputStream fileInputStream = null;
		FileOutputStream fileOutputStream = null;
		try {
			fileInputStream = new FileInputStream(pathXLSFile);
			xssfWorkbook = new XSSFWorkbook(fileInputStream);

			if (!isSheetExist(sheetName)) {
				return false;
			}

			int index = xssfWorkbook.getSheetIndex(sheetName);
			XSSFCellStyle style = xssfWorkbook.createCellStyle();
			style.setFillForegroundColor(HSSFColor.HSSFColorPredefined.BRIGHT_GREEN.getIndex());
			style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			XSSFSheet sheet = xssfWorkbook.getSheetAt(index);

			XSSFRow row = sheet.getRow(0);
			if (row == null)
				row = sheet.createRow(0);

			XSSFCell cell;
			if (row.getLastCellNum() == -1)
				cell = row.createCell(0);
			else
				cell = row.createCell(row.getLastCellNum());

			cell.setCellValue(colName);
			cell.setCellStyle(style);

			fileOutputStream = new FileOutputStream(pathXLSFile);
			xssfWorkbook.write(fileOutputStream);

		} catch (Exception e) {
			e.printStackTrace();
			return false;
		} finally {

			try {
				if (fileOutputStream != null) {
					fileOutputStream.close();
				}
				if (fileInputStream != null) {
					fileInputStream.close();
				}
			} catch (IOException e) {
				e.printStackTrace();
			}
		}

		return true;
	}

	/**
	 * removes a column and all the contents
	 * 
	 * @param sheetName
	 * @param colNum
	 * @return
	 */
	public boolean removeColumn(String sheetName, int colNum) {

		FileInputStream fileInputStream = null;
		FileOutputStream fileOutputStream = null;
		try {
			if (!isSheetExist(sheetName))
				return false;
			fileInputStream = new FileInputStream(pathXLSFile);
			xssfWorkbook = new XSSFWorkbook(fileInputStream);
			XSSFSheet sheet = xssfWorkbook.getSheet(sheetName);
			XSSFCellStyle style = xssfWorkbook.createCellStyle();
			style.setFillForegroundColor(HSSFColor.HSSFColorPredefined.BRIGHT_GREEN.getIndex());
			style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			for (int i = 0; i < getRowCount(sheetName); i++) {
				XSSFRow row = sheet.getRow(i);
				if (row != null) {
					XSSFCell cell = row.getCell(colNum - 1);
					if (cell != null) {
						cell.setCellStyle(style);
						row.removeCell(cell);
					}
				}
			}
			fileOutputStream = new FileOutputStream(pathXLSFile);
			xssfWorkbook.write(fileOutputStream);
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		} finally {
			try {
				if (fileOutputStream != null) {
					fileOutputStream.close();
				}
				if (fileInputStream != null) {
					fileInputStream.close();
				}
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		return true;
	}

	/**
	 * String sheetName, String testCaseName,String keyword ,String URL,String
	 * message
	 * 
	 * @param sheetName
	 * @param screenShotColName
	 * @param testCaseName
	 * @param index
	 * @param url
	 * @param message
	 * @return
	 */
	public boolean addHyperLink(String sheetName, String screenShotColName, String testCaseName, int index, String url,
			String message) {

		url = url.replace('\\', '/');
		if (!isSheetExist(sheetName))
			return false;

		for (int i = 2; i <= getRowCount(sheetName); i++) {
			if (getCellData(sheetName, 0, i).equalsIgnoreCase(testCaseName)) {
				setCellData(sheetName, screenShotColName, i + index, message, url);
				break;
			}
		}

		return true;
	}

	/**
	 * returns true if data is set successfully else false
	 * 
	 * @param sheetName
	 * @param colName
	 * @param rowNum
	 * @param data
	 * @return
	 */
	public boolean setCellData(String sheetName, String colName, int rowNum, String data) {

		FileInputStream fileInputStream = null;
		FileOutputStream fileOutputStream = null;
		try {
			fileInputStream = new FileInputStream(pathXLSFile);
			xssfWorkbook = new XSSFWorkbook(fileInputStream);

			if (rowNum <= 0)
				return false;

			int index = xssfWorkbook.getSheetIndex(sheetName);
			int colNum = -1;
			if (index == -1)
				return false;

			XSSFSheet sheet = xssfWorkbook.getSheetAt(index);

			XSSFRow row = sheet.getRow(0);
			for (int i = 0; i < row.getLastCellNum(); i++) {

				if (row.getCell(i).getStringCellValue().trim().equals(colName))
					colNum = i;
			}
			if (colNum == -1)
				return false;

			sheet.autoSizeColumn(colNum);
			row = sheet.getRow(rowNum - 1);
			if (row == null)
				row = sheet.createRow(rowNum - 1);

			XSSFCell cell = row.getCell(colNum);
			if (cell == null)
				cell = row.createCell(colNum);

			cell.setCellValue(data);
			fileOutputStream = new FileOutputStream(pathXLSFile);
			xssfWorkbook.write(fileOutputStream);

		} catch (Exception e) {
			e.printStackTrace();
			return false;
		} finally {
			try {
				if (fileOutputStream != null) {
					fileOutputStream.close();
				}
				if (fileInputStream != null) {
					fileInputStream.close();
				}
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		return true;
	}

	/**
	 * returns true if data is set successfully else false
	 * 
	 * @param sheetName
	 * @param colName
	 * @param rowNum
	 * @param data
	 * @param url
	 * @return
	 */
	public boolean setCellData(String sheetName, String colName, int rowNum, String data, String url) {

		FileInputStream fileInputStream = null;
		FileOutputStream fileOutputStream = null;
		try {
			fileInputStream = new FileInputStream(pathXLSFile);
			xssfWorkbook = new XSSFWorkbook(fileInputStream);

			if (rowNum <= 0)
				return false;

			int index = xssfWorkbook.getSheetIndex(sheetName);
			int colNum = -1;
			if (index == -1)
				return false;

			XSSFSheet sheet = xssfWorkbook.getSheetAt(index);

			XSSFRow row = sheet.getRow(0);
			for (int i = 0; i < row.getLastCellNum(); i++) {

				if (row.getCell(i).getStringCellValue().trim().equalsIgnoreCase(colName))
					colNum = i;
			}

			if (colNum == -1)
				return false;
			sheet.autoSizeColumn(colNum); // ashish
			row = sheet.getRow(rowNum - 1);
			if (row == null)
				row = sheet.createRow(rowNum - 1);

			XSSFCell cell = row.getCell(colNum);
			if (cell == null)
				cell = row.createCell(colNum);

			cell.setCellValue(data);
			XSSFCreationHelper createHelper = xssfWorkbook.getCreationHelper();

			// cell style for hyperlinks
			// by default hypelrinks are blue and underlined
			CellStyle hlink_style = xssfWorkbook.createCellStyle();
			XSSFFont hlink_font = xssfWorkbook.createFont();
			hlink_font.setUnderline(XSSFFont.U_SINGLE);
			hlink_font.setColor(IndexedColors.BLUE.getIndex());
			hlink_style.setFont(hlink_font);
			// hlink_style.setWrapText(true);

			XSSFHyperlink link = createHelper.createHyperlink(HyperlinkType.FILE);
			link.setAddress(url);
			cell.setHyperlink(link);
			cell.setCellStyle(hlink_style);

			fileOutputStream = new FileOutputStream(pathXLSFile);
			xssfWorkbook.write(fileOutputStream);

		} catch (Exception e) {
			e.printStackTrace();
			return false;
		} finally {
			try {
				if (fileOutputStream != null) {
					fileOutputStream.close();
				}
				if (fileInputStream != null) {
					fileInputStream.close();
				}
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		return true;
	}

	/**
	 * 
	 * @param cell
	 * @return
	 */
	private String getCellValue(XSSFCell cell) {
		if (cell.getCellTypeEnum() == CellType.STRING) {
			return cell.getStringCellValue();

		} else if (cell.getCellTypeEnum() == CellType.NUMERIC || cell.getCellTypeEnum() == CellType.FORMULA) {
			String cellText = String.valueOf(cell.getNumericCellValue());
			if (HSSFDateUtil.isCellDateFormatted(cell)) {
				DateFormat df = new SimpleDateFormat("dd/MM/yyyy");
				Date date = cell.getDateCellValue();
				cellText = df.format(date);
			}

			return cellText;

		} else if (cell.getCellTypeEnum() == CellType.BLANK) {
			return "";
		} else {
			return String.valueOf(cell.getBooleanCellValue());
		}
	}

}
