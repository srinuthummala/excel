package excel;

import java.io.FileInputStream;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadDatFromExcel {

	public static void main(String[] args) throws Exception {

		try {
			FileInputStream file = new FileInputStream("C:/Users/CHOWDARY/Desktop/sample.xlsx");

			XSSFWorkbook workbook = new XSSFWorkbook(file);
			XSSFSheet sheet = workbook.getSheet("Sheet0");

			System.out.println(sheet.getRow(0).getCell(0).getStringCellValue());

			int rowcount = sheet.getLastRowNum();
			System.out.println(rowcount);

			for (int i = 0; i <= rowcount; i++) {
				Row row = sheet.getRow(i);

				for (int j = 0; j < row.getLastCellNum(); j++) {
					// String cellvalue = rowvalue.getCell(j).getStringCellValue();
					// System.out.println(rowvalue.getCell(j).getStringCellValue());
					Cell cell = row.getCell(j);
					if (cell.getCellTypeEnum() == CellType.STRING) {
						System.out.println(cell.getStringCellValue());
					} else if (cell.getCellTypeEnum() == CellType.NUMERIC
							|| cell.getCellTypeEnum() == CellType.FORMULA) {
						System.out.println(cell.getNumericCellValue());
						if (HSSFDateUtil.isCellDateFormatted(cell)) {
							SimpleDateFormat sdf = new SimpleDateFormat("dd/mm/yy");
							Date date = cell.getDateCellValue();
							System.out.println(sdf.format(date));
						}
					} else if (cell.getCellTypeEnum() == CellType.BLANK)
						System.out.println("");
				}
			}
			workbook.close();
		} catch (Exception e) {
			System.out.println(e);
		}

	}

}
