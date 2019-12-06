package utils;

public class ExcelUtilsImplementation {
	
	public static void main(String[] args) {
		String projectPath = System.getProperty("user.dir");
		ExcelUtils obj = new ExcelUtils(projectPath +"/Excel/Book1.xlsx" , "LoginCredentials");
		
		ExcelUtils.getRowCount();
		ExcelUtils.getColumnCount();
		ExcelUtils.getCellDataString(0, 0);
		ExcelUtils.getCellNumericData(1, 1);
	}
	
	
	
}
