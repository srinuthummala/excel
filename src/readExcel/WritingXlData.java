package readExcel;

public class WritingXlData {

	public static void main(String[] args) throws Exception {
		WriteDataToXlByCellNum xlWrite = new WriteDataToXlByCellNum("C:\\Users\\CHOWDARY\\Desktop\\xlfile.xlsx");

		xlWrite.setCellData("sri", 2, 4, "fail");

	}

}
