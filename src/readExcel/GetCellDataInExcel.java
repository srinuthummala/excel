package readExcel;

import java.io.FileInputStream;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class GetCellDataInExcel {

	public FileInputStream fis = null;
	public XSSFWorkbook workbook = null;
	public XSSFSheet sheet = null;
	public XSSFRow row = null;
	public XSSFCell cell = null;
	public String cellValue;

	public GetCellDataInExcel(String xlFilePath) throws Exception {

		fis = new FileInputStream(xlFilePath);
		workbook = new XSSFWorkbook(fis);

		fis.close();
	}

	public String getCellData(String sheetName, int rowNum, int colNum) {

		try {
			sheet = workbook.getSheet(sheetName);
			row = sheet.getRow(rowNum);
			cell = row.getCell(colNum);

			if (cell.getCellTypeEnum() == CellType.STRING) {

				cellValue = cell.getStringCellValue();

			} else if (cell.getCellTypeEnum() == CellType.NUMERIC || cell.getCellTypeEnum() == CellType.FORMULA) {
				cellValue = String.valueOf(cell.getNumericCellValue());

				if (HSSFDateUtil.isCellDateFormatted(cell)) {
					SimpleDateFormat df = new SimpleDateFormat("dd/mm/yy");
					Date date = cell.getDateCellValue();
					cellValue = df.format(date);
				}

			} else if (cell.getCellTypeEnum() == CellType.BLANK)

				cellValue = "";
			else if (cell.getCellTypeEnum() == CellType.BOOLEAN)

				cellValue = String.valueOf(cell.getBooleanCellValue());
		} catch (Exception e) {

			e.printStackTrace();

			cellValue = "no matched value";
		}
		return cellValue;

	}
}
