package utils;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtils {
	static String projectPath;
	static XSSFWorkbook workbook;
	static XSSFSheet sheet;

	public ExcelUtils(String excelPath, String sheetName) {
		try {
			
			workbook = new XSSFWorkbook(excelPath);
			sheet = workbook.getSheet(sheetName);
		} catch (Exception e) {
			e.getStackTrace();
		}
	}

	

	public static int getRowCount() {
		int rowCount = 0;
		try {

			rowCount = sheet.getPhysicalNumberOfRows();
			System.out.println("No.of rows: " + rowCount);

		} catch (Exception exp) {
			System.out.println(exp.getCause());

			System.out.println(exp.getMessage());
			exp.getStackTrace();
		}
		return rowCount;
	}
	
	public static int getColumnCount() {
		int colCount = 0;
		try {

			colCount = sheet.getRow(0).getPhysicalNumberOfCells();
			System.out.println("No.of columns: " + colCount);

		} catch (Exception exp) {
			System.out.println(exp.getCause());

			System.out.println(exp.getMessage());
			exp.getStackTrace();
		}
		return colCount;
	}

	public static String getCellDataString(int rownum, int colnum) {
		 System.out.println(sheet.getRow(rownum).getCell(colnum).getCellType());
		String cellData = null;
		try {

			cellData = sheet.getRow(rownum).getCell(colnum).getStringCellValue();
			//System.out.println("Cell data: " + cellData);

		} catch (Exception exp) {
			System.out.println(exp.getCause());

			System.out.println(exp.getMessage());
			exp.getStackTrace();
		}
		   return cellData;
	}

	public static void getCellNumericData(int rownum, int colnum) {
		try {

			Double cellDataNumber = sheet.getRow(rownum).getCell(colnum).getNumericCellValue();
			//System.out.println("Cell data: " + cellDataNumber);

		} catch (Exception exp) {
			System.out.println(exp.getCause());

			System.out.println(exp.getMessage());
			exp.getStackTrace();
		}

	}

}
