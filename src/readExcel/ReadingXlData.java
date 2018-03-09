package readExcel;

public class ReadingXlData {

	public static void main(String[] args) throws Exception {
		/*
		 * GetCellDataInExcel xlData = new
		 * GetCellDataInExcel("C:\\Users\\CHOWDARY\\Desktop\\xlfile.xlsx");
		 * System.out.println(xlData.getCellData("sri", 0, 0));
		 * System.out.println(xlData.getCellData("sri", 0, 1));
		 * 
		 * System.out.println(xlData.getCellData("sri", 1, 0));
		 * 
		 * System.out.println(xlData.getCellData("sri", 1, 1));
		 * System.out.println(xlData.getCellData("sri", 1, 2));
		 * System.out.println(xlData.getCellData("sri", 1, 3));
		 * System.out.println(xlData.getCellData("sri", 2, 3));
		 */

		GetCellDataByColName xlData = new GetCellDataByColName("C:\\Users\\CHOWDARY\\Desktop\\xlfile.xlsx");
		System.out.println(xlData.cellData("sri", "userName", 1));

		System.out.println(xlData.cellData("sri", "userName", 4));

	}

}
