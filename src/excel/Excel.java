package excel;

import java.io.File;
import java.io.FileOutputStream;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel {

	public static void main(String[] args) throws Exception {
		//comment
		// File file = new File("C:/Users/CHOWDARY/Desktop/sample.xlsx");
		// file.createNewFile();
		try {
			XSSFWorkbook workbook = new XSSFWorkbook();
			FileOutputStream file = new FileOutputStream(new File("C:/Users/CHOWDARY/Desktop/sample.xlsx"));
			@SuppressWarnings("unused")
			XSSFSheet worksheet = workbook.createSheet();
			workbook.write(file);
			file.close();
			workbook.close();
		} catch (Exception e) {
			System.out.println(e);
		}
		System.out.println("excel file is created");
	}

}
