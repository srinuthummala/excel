package readExcel;

public class ReadWriteXl {

	public static void main(String[] args) throws Exception {

		/*
		 * XlReadWriteAPI GetCellData = new
		 * XlReadWriteAPI("C:\\Users\\CHOWDARY\\Desktop\\xlfile.xlsx");
		 * System.out.println(GetCellData.readXlWithColNum("sri", 1, 0));
		 * System.out.println(GetCellData.readXlWithColNum("sri", 1, 1));
		 * System.out.println(GetCellData.readXlWithColNum("sri", 1, 2));
		 * System.out.println(GetCellData.readXlWithColNum("sri", 1, 3));
		 * System.out.println(GetCellData.readXlWithColNum("sri", 1, 4));
		 * System.out.println(GetCellData.readXlWithColNum("sri", 1, 5));
		 * System.out.println(GetCellData.readXlWithColNum("sri", 2, 1));
		 * 
		 * 
		 * XlReadWriteAPI GetCellData1 = new
		 * XlReadWriteAPI("C:\\Users\\CHOWDARY\\Desktop\\xlfile.xlsx");
		 * System.out.println(GetCellData1.readXlWithCellName("sri", "password", 3));
		 * System.out.println(GetCellData1.readXlWithCellName("sri", "userName", 1));
		 * 
		 * 
		 * XlReadWriteAPI setCellData = new
		 * XlReadWriteAPI("C:\\Users\\CHOWDARY\\Desktop\\xlfile.xlsx");
		 * System.out.println(setCellData.writeXlWithCellNum("sri", 1, 5, "fail"));
		 * 
		 * XlReadWriteAPI setCellData1 = new
		 * XlReadWriteAPI("C:\\Users\\CHOWDARY\\Desktop\\xlfile.xlsx");
		 * System.out.println(setCellData1.writeXlWithCellName("sri", 2, "result",
		 * "true"));
		 */

		XlReadWriteAPI rowCount = new XlReadWriteAPI("C:\\Users\\CHOWDARY\\Desktop\\xlfile.xlsx");
		rowCount.rowColumnCount();

	}

}
