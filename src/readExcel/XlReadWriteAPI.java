package readExcel;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XlReadWriteAPI {

	public FileInputStream fis = null;
	public FileOutputStream fos = null;
	public XSSFWorkbook workbook = null;
	public XSSFSheet sheet = null;
	public XSSFRow row = null;
	public XSSFCell cell = null;
	public String cellRetunData;
	String cellValue;
	String xlPath;

	public XlReadWriteAPI(String xlPath) throws Exception {
		this.xlPath = xlPath;
		fis = new FileInputStream(xlPath);
		workbook = new XSSFWorkbook(fis);

		fis.close();
	}

	// Reading excel file data using the cell number and row number
	public String readXlWithColNum(String sheetName, int rowNum, int colNum) {

		try {

			sheet = workbook.getSheet(sheetName);
			row = sheet.getRow(rowNum);
			cell = row.getCell(colNum);

			if (cell.getCellTypeEnum() == CellType.STRING) {
				cellRetunData = cell.getStringCellValue();
			} else if (cell.getCellTypeEnum() == CellType.NUMERIC || cell.getCellTypeEnum() == CellType.FORMULA) {
				cellRetunData = String.valueOf(cell.getNumericCellValue());
				if (HSSFDateUtil.isCellDateFormatted(cell)) {
					SimpleDateFormat sdf = new SimpleDateFormat("dd/mm/yy");
					Date date = cell.getDateCellValue();
					cellRetunData = sdf.format(date);
				}
			} else if (cell.getCellTypeEnum() == CellType.BLANK)
				cellRetunData = "";
			else if (cell.getCellTypeEnum() == CellType.BOOLEAN)
				cellRetunData = String.valueOf(cell.getBooleanCellValue());

		} catch (Exception e) {
			e.printStackTrace();
			cellRetunData = "cell formate Exception";
		}
		return cellRetunData;
	}

	// Reading excel file data using the cell name and row number
	public String readXlWithCellName(String sheetName, String colName, int rowNum) {

		try {
			int colNum = -1;
			sheet = workbook.getSheet(sheetName);
			row = sheet.getRow(0);
			for (int i = 0; i < row.getLastCellNum(); i++) {
				if (row.getCell(i).getStringCellValue().trim().equals(colName))
					colNum = i;
			}

			row = sheet.getRow(rowNum);
			cell = row.getCell(colNum);

			if (cell.getCellTypeEnum() == CellType.STRING) {
				cellValue = cell.getStringCellValue();
			} else if (cell.getCellTypeEnum() == CellType.NUMERIC || cell.getCellTypeEnum() == CellType.FORMULA) {
				cellValue = String.valueOf(cell.getNumericCellValue());
				if (HSSFDateUtil.isCellDateFormatted(cell)) {
					SimpleDateFormat sdf = new SimpleDateFormat("dd/mm/yy");
					Date date = cell.getDateCellValue();
					cellValue = sdf.format(date);
				}
			} else if (cell.getCellTypeEnum() == CellType.BLANK)
				cellValue = "";
			else if (cell.getCellTypeEnum() == CellType.BOOLEAN)
				cellValue = String.valueOf(cell.getBooleanCellValue());
		} catch (Exception e) {

			cellValue = "cell format Exception";
		}
		return cellValue;
	}

	// Reading excel file data using the cell number and row number
	public Boolean writeXlWithCellNum(String sheetName, int rowNum, int colNum, String value) {
		try {
			sheet = workbook.getSheet(sheetName);

			row = sheet.getRow(rowNum);
			if (row == null)
				row = sheet.createRow(rowNum);
			cell = row.getCell(colNum);
			if (cell == null)
				cell = row.createCell(colNum);

			cell.setCellValue(value);

			fos = new FileOutputStream(xlPath);
			workbook.write(fos);

		} catch (Exception e) {
			return false;
		}
		return true;
	}

	// Reading excel file data using the cell name and row number
	public Boolean writeXlWithCellName(String sheetName, int rowNum, String colName, String value) {
		try {
			int colNum = -1;
			sheet = workbook.getSheet(sheetName);
			row = sheet.getRow(0);
			for (int i = 0; i < row.getLastCellNum(); i++) {
				if (row.getCell(i).getStringCellValue().trim().equals(colName))
					colNum = i;
			}

			row = sheet.getRow(rowNum);
			if (row == null)
				row = sheet.createRow(rowNum);
			cell = row.getCell(colNum);
			if (cell == null)
				cell = row.createCell(colNum);

			cell.setCellValue(value);

			fos = new FileOutputStream(xlPath);
			workbook.write(fos);

		} catch (Exception e) {
			return false;
		}
		return true;
	}

	// row count and column count

	public void rowColumnCount() {
		sheet = workbook.getSheetAt(0);
		int rowCount = sheet.getLastRowNum();

		int colCount = sheet.getRow(0).getLastCellNum();

		System.out.println(rowCount);
		System.out.println(colCount);
	}
}
