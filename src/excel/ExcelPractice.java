package excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelPractice {

	public static void main(String[] args) throws Exception {
		/*
		 * ExcelPractice ep = new ExcelPractice(); try { ep.excelcreate(); } catch
		 * (Exception e) {
		 * 
		 * e.printStackTrace();
		 */
		// }

		try {
			ExcelPractice.xlWrite();
		} catch (Exception e) {
			System.out.println("e");
		}

		/*
		 * try { ExcelPractice.xlRead(); } catch (Exception e) { // TODO Auto-generated
		 * catch block e.printStackTrace(); }
		 */
	}

	public void excelcreate() throws Exception {

		File xlfile = new File("C:/Users/CHOWDARY/Desktop/xlfile1.xlsx");
		boolean bo = xlfile.exists();
		if (bo == true) {
			// xlfile.delete();
			System.out.println("file already exists");
			bo = false;

		} else {
			try {
				XSSFWorkbook wb = new XSSFWorkbook();
				FileOutputStream fos = new FileOutputStream(xlfile);

				@SuppressWarnings("unused")
				XSSFSheet sheet = wb.createSheet();
				wb.write(fos);

				wb.close();
			} catch (Exception e) {
				System.out.println("fileexception");
			}
		}

	}

	public static void xlWrite() throws Exception {
		File xlFile = new File("C:/Users/CHOWDARY/Desktop/xlfile.xlsx");

		try {
			XSSFWorkbook wb = new XSSFWorkbook();
			XSSFSheet sheet = wb.createSheet("sri");

			for (int i = 0, k = 0; i < 5; i++, k++) {
				Row row = sheet.createRow(i);
				int j = 0;

				int Age = 25;
				String Name = "srinu";

				row.createCell(j).setCellValue(Name + k);
				row.createCell(j + 1).setCellValue(Age + k);

				FileOutputStream fos = new FileOutputStream(xlFile);
				wb.write(fos);

			}

			/*
			 * sheet.createRow(0).createCell(0).setCellValue("name");
			 * sheet.getRow(0).createCell(1).setCellValue("vars");
			 * sheet.createRow(1).createCell(0).setCellValue("srinu");
			 * sheet.getRow(1).createCell(1).setCellValue("some");
			 * sheet.createRow(2).createCell(0).setCellValue(true); FileOutputStream fos =
			 * new FileOutputStream(xlFile);
			 */
			// wb.write(fos);

			wb.close();

		} catch (Exception e) {
			e.printStackTrace();

		}

	}

	public static void xlRead() throws Exception {
		File xlFile = new File("C:/Users/CHOWDARY/Desktop/xlfile.xlsx");
		try {
			FileInputStream fis = new FileInputStream(xlFile);
			XSSFWorkbook wb = new XSSFWorkbook(fis);
			XSSFSheet sheet = wb.getSheet("sri");

			int rowcount = sheet.getLastRowNum();

			System.out.println(rowcount);

			for (int i = 0; i <= rowcount; i++) {
				Row rowvalue = sheet.getRow(i);
				// System.out.println(rowvalue);
				for (int j = 0; j < rowvalue.getLastCellNum(); j++) {
					System.out.println(rowvalue.getCell(j).getStringCellValue());
				}
			}
			wb.close();

		} catch (Exception e) {
			e.printStackTrace();
		}

	}

}
