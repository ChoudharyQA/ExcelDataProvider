package DataDriven.ExcelDataProvider;

import java.io.FileInputStream;

import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class DataProvide {

	// Multiple sets of data to your test
	//Through Array we will send the data
	// 5 sets of data as 5 arrays from data provider to your test
	// Then your test will run 5 times with 5 different sets of data as Arrays
	
	
	DataFormatter formatter = new DataFormatter();
	@Test(dataProvider="driveTest")
	public void testCaseData(String Greeting, String Communication, String ID) {
		
		System.out.println(Greeting+Communication+ID);
		
	}
	
	@DataProvider(name="driveTest")
	public Object[][] getData() throws IOException {
		//Object[][] data = {{"hello","text","1"},{"bye","message","143"},{"solo","call","453"}};
		//return data;
		
		
		//Now will catch the data from excel and then send to an array each and then those stored value in array should be send to our test cases 
		//Every raw of the excel should be send to 1 array each
		
		FileInputStream fis = new FileInputStream("C:\\Users\\admin\\Desktop\\Excel2.xlsx");
		XSSFWorkbook wb =new XSSFWorkbook(fis);
		XSSFSheet sheet = wb.getSheetAt(0);
		int rowCount = sheet.getPhysicalNumberOfRows();
		XSSFRow row = sheet.getRow(0);
		int columnCount = row.getLastCellNum();
		Object data[][] = new Object [rowCount-1][columnCount];
		for(int i=0;i<rowCount-1;i++) {
			
			row = sheet.getRow(i+1);
			for(int j=0;j<columnCount;j++) {
				
				
				
				XSSFCell cell = row.getCell(j);
				data  [i][j]= formatter.formatCellValue(cell);
				
			}
			
		}
		
		
		return data;
		
		
		
		
		
	}
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
}
