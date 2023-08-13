
package Demo;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

/**
 * @author santhosh.naik
 * Mobile: +919071076387
 * Email: snaik.santhosh@gmail.com
 */
public class Read_Write_Excel {

	//main method
	public static void main(String[] args) throws Exception
	{
		
		 //Get the excel file and create an input stream for excel
		 FileInputStream fis = new FileInputStream("C:\\Users\\srinivasan_k\\stsworkspace\\Sample_Project\\results\\Age_Validation.xlsx");
		 
		 //load the input stream to a workbook object
		 //Use XSSF for (.xlsx) excel file and HSSF for (.xls) excel file
		 XSSFWorkbook wb = new XSSFWorkbook(fis);
		 
		 //get the sheet from the workbook by index
		 XSSFSheet sheet = wb.getSheet("Age");
		 
		 //Count the total number of rows present in the sheet
		 int rowcount = sheet.getLastRowNum();
		 System.out.println(" Total number of rows present in the sheet : "+rowcount);
		 
		 //get column count present in the sheet
		 int colcount = sheet.getRow(1).getLastCellNum();
		 System.out.println(" Total number of columns present in the sheet : "+colcount);
		 
		 //get the data from sheet by iterating through cells
		 //by using for loop
		 for(int i = 1; i<=rowcount; i++)
		  {
			 XSSFCell cell = sheet.getRow(i).getCell(1);
			 String celltext="";
			 
			 //Get celltype values
			 if(cell.getCellType()==Cell.CELL_TYPE_STRING)
			 {
				 celltext=cell.getStringCellValue();
			 }
			 else if(cell.getCellType()==Cell.CELL_TYPE_NUMERIC)
			 {
				  celltext=String.valueOf(cell.getNumericCellValue());
			 }
			 else if(cell.getCellType()==Cell.CELL_TYPE_BLANK)
			 {
				 celltext="";
			 }
		  
		  //Check the age and set the Cell value into excel
			 if(Double.parseDouble(celltext)>=18)
			 {
				 sheet.getRow(i).getCell(2).setCellValue("Major");
			 }
			 else
			 {
				 sheet.getRow(i).getCell(2).setCellValue("Minor");
			 }
			 
		  }//End of for loop
		 
		 //close the file input stream
		 fis.close();
		 

	//Open an excel to write the data into workbook
	FileOutputStream fos = new FileOutputStream("C:\\Users\\srinivasan_k\\stsworkspace\\Sample_Project\\results\\Age_Validation_Results.xlsx");
	
	//Write into workbook
	wb.write(fos);
	
	//close fileoutstream
	fos.close();

	}

}
