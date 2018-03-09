package excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Xl {

	public static void main(String[] args) throws Exception {
		// Xl.createXl();
		// Xl.xlWrite();
		Xl.xlRead();

	}

	public static void createXl() throws FileNotFoundException, IOException {
		File f = new File("xl.xlsx");
		XSSFWorkbook wb = new XSSFWorkbook();
		FileOutputStream fos = new FileOutputStream(f);
		XSSFSheet sheet = wb.createSheet();
		wb.write(fos);

	}

	public static void xlWrite() throws IOException {

		FileOutputStream fos = new FileOutputStream(new File("xl.xlsx"));

		XSSFWorkbook wb = new XSSFWorkbook();
		XSSFSheet sheet = wb.createSheet();

		sheet.createRow(0).createCell(0).setCellValue("student");
		sheet.createRow(1).createCell(0).setCellValue("srinu");
		sheet.createRow(2).createCell(0).setCellValue("srinu");
		sheet.createRow(3).createCell(0).setCellValue("srinu");

		wb.write(fos);

		wb.close();

	}

	public static void xlRead() throws Exception {

		try {
			FileInputStream fis = new FileInputStream(new File("xl.xlsx"));

			XSSFWorkbook wb = new XSSFWorkbook(fis);
			XSSFSheet sheet = wb.getSheetAt(0);

			int rowcount = sheet.getLastRowNum();
			System.out.println(rowcount);

			for (int i = 0; i < rowcount; i++) {
				Row rowvalue = sheet.getRow(i);
				// System.out.println(rowvalue);

				for (int j = 0; j < rowvalue.getLastCellNum(); j++) {
					System.out.println(rowvalue.getCell(j).getStringCellValue());
				}
			}
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

}
