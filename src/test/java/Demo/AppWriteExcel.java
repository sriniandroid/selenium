
package Demo;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.common.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFCreationHelper;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONObject;

public class AppWriteExcel {

	// main method
	public static void main(String[] args) throws Exception {

		// Get the excel file and create an input stream for excel
		FileInputStream fis = new FileInputStream(
				"C:\\Users\\srinivasan_k\\stsworkspace\\Sample_Project\\results\\App_Data.xlsx");

		// Open an excel to write the data into workbook
		FileOutputStream fos = new FileOutputStream(
				"C:\\Users\\srinivasan_k\\stsworkspace\\Sample_Project\\results\\App_Data_Results.xlsx");

		
		// load the input stream to a workbook object
		// Use XSSF for (.xlsx) excel file and HSSF for (.xls) excel file
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		

		createResults(wb);

		// get the sheet from the workbook by index
		XSSFSheet sheet = wb.getSheet("Data");

		// Count the total number of rows present in the sheet
		int rowcount = sheet.getLastRowNum();
		System.out.println(" Total number of rows present in the sheet : " + rowcount);

		// get the data from sheet by iterating through cells
		// by using for loop
		for (int i = 0; i <= rowcount; i += 2) {

			String reqKey = "";
			String resKey = "";

			XSSFCell cell = sheet.getRow(i).getCell(0);
			if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
				reqKey = cell.getStringCellValue();
				reqKey = getKey(reqKey);
			}
			cell = sheet.getRow(i + 1).getCell(0);
			if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
				resKey = cell.getStringCellValue();
				resKey = getKey(resKey);
			}

			if (reqKey.equals(resKey)) {
				appendResults(wb, resKey, resKey, "Match", i);
				System.out.println("Keys Matching " + resKey);
			} else {
				appendResults(wb, resKey, resKey, "Not Match", i);
				System.out.println("Keys Not Matching " + resKey);
			}

		} // End of for loop

		// close the file input stream
		fis.close();

		// Write into workbook
		wb.write(fos);

		// close fileoutstream
		fos.close();

	}

	public static void appendResults(XSSFWorkbook wb, String requestKey, String responseKey, String result, int fileRowNumber) {
		XSSFSheet results = wb.getSheet("Test Results");
		int rowcount = results.getLastRowNum();
		
		results.createRow(++rowcount);
		  
		results.getRow(rowcount).createCell(0).setCellValue(requestKey);
		results.getRow(rowcount).createCell(1).setCellValue(responseKey);
		results.getRow(rowcount).createCell(2).setCellValue(result);
		results.getRow(rowcount).createCell(3).setCellValue("");
		
		Cell linkCell = results.getRow(rowcount).createCell(3);

		XSSFCellStyle style = wb.createCellStyle();
		XSSFFont font = wb.createFont();
		font.setFontHeightInPoints((short) 10);
		font.setUnderline(Font.U_SINGLE);
		font.setColor(IndexedColors.BLUE.getIndex());
		style.setFont(font);
		
		XSSFCreationHelper createHelper = wb.getCreationHelper();
		linkCell.setCellValue("File Link");
		XSSFHyperlink link = createHelper.createHyperlink(Hyperlink.LINK_DOCUMENT);
		link.setAddress("App_Data_Results.xlsx"); 
		linkCell.setHyperlink(link);
		
		linkCell.setCellStyle(style);
	}
	
	public static void createResults(XSSFWorkbook wb) {
		XSSFSheet results = wb.createSheet("Test Results");
		results.createRow(0);
		results.getRow(0).createCell(0).setCellValue("Request Key");
		results.getRow(0).createCell(1).setCellValue("Response Key");
		results.getRow(0).createCell(2).setCellValue("Result");
		results.getRow(0).createCell(3).setCellValue("Link");
		
		XSSFFont font = wb.createFont();
		font.setFontHeightInPoints((short) 10);
		font.setColor(IndexedColors.WHITE.getIndex());
		font.setBold(true);
		font.setItalic(false);
		
		XSSFCellStyle style = wb.createCellStyle();
		style.setFillForegroundColor(IndexedColors.DARK_BLUE.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setAlignment(CellStyle.ALIGN_CENTER);
		style.setFont(font);
		
		results.getRow(0).getCell(0).setCellStyle(style);
		results.getRow(0).getCell(1).setCellStyle(style);
		results.getRow(0).getCell(2).setCellStyle(style);
		results.getRow(0).getCell(3).setCellStyle(style);
		
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
