package utils;

public class ExcelDataProvider {
	
	public static void main(String[] args) {
		String projectPath = System.getProperty("user.dir");
		testData(projectPath +"/Excel/Book1.xlsx" , "LoginCredentials");
		
	}
	
	public static Object[][] testData(String excelPath, String sheetName) {
		ExcelUtils excel = new ExcelUtils(excelPath, sheetName);
		
		int rowCount = excel.getRowCount();
		int colCount = excel.getColumnCount();
		
		Object data[][] = new Object[rowCount-1][colCount];
		
		for(int i=1; i<rowCount; i++) {
			for(int j=0; j<colCount; j++) {
				String cellData = excel.getCellDataString(i, j);
				System.out.print(cellData + " | ");
				data[i-1][j] = cellData;
			}
			System.out.println();
		}
			return data;
			
		
		
		
	}

}
