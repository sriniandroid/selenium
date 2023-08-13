
package Demo;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONObject;

public class AppWriteExcel {

	// main method
	public static void main(String[] args) throws Exception {

		// Get the excel file and create an input stream for excel
		FileInputStream fis = new FileInputStream(
				"C:\\Users\\srinivasan_k\\stsworkspace\\Sample_Project\\results\\App_Data.xlsx");

		//Open an excel to write the data into workbook
		FileOutputStream fos = new FileOutputStream("C:\\Users\\srinivasan_k\\stsworkspace\\Sample_Project\\results\\Age_Validation_Results.xlsx");
				
		// load the input stream to a workbook object
		// Use XSSF for (.xlsx) excel file and HSSF for (.xls) excel file
		XSSFWorkbook wb = new XSSFWorkbook(fis);

		// get the sheet from the workbook by index
		XSSFSheet sheet = wb.getSheet("Age");

		// Count the total number of rows present in the sheet
		int rowcount = sheet.getLastRowNum();
		System.out.println(" Total number of rows present in the sheet : " + rowcount); 

		// get the data from sheet by iterating through cells
		// by using for loop
		for (int i = 0; i <= rowcount; i+=2) {
			
			String reqKey = "";
			String resKey = "";
			
			XSSFCell cell = sheet.getRow(i).getCell(0);
			if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
				reqKey = cell.getStringCellValue();
				reqKey = getKey(reqKey);
			} 
			cell = sheet.getRow(i+1).getCell(0);
			if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
				resKey = cell.getStringCellValue();
				resKey = getKey(resKey);
			}
			
			if(reqKey.equals(resKey)) {
				System.out.println("Keys Matching "+ resKey);
			}else {
				System.out.println("Keys Not Matching "+ resKey);
			}

		} // End of for loop

		// close the file input stream
		fis.close();
		
		
		//Write into workbook
		wb.write(fos);
		
		//close fileoutstream
		fos.close();

	}

	public static String getKey(String data) {
		if (data == null) {
			return "";
		}
		JSONObject jsonObject = new JSONObject(data);
		if (jsonObject.has("request")) {
			return jsonObject.get("request").toString();
		} else if (jsonObject.has("response")) {
			return jsonObject.get("response").toString();
		}
		return "0";
	}

}
