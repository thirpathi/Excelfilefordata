package excelfileforread;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelFileForRead {

	  public static void main(String[] args) throws IOException, EncryptedDocumentException, InvalidFormatException {

	   FileInputStream fip = new FileInputStream("D:\\rameshsoft\\workspace\\pb8batch1\\src\\com\\rameshsoft\\automation\\testdata\\TestData.xls");
	   Workbook workbook = WorkbookFactory.create(fip);
	    
	   // go to sheet name (i.e getsheet)
	   Sheet sheet1 = workbook.getSheet("ramesh");
	    
	   // select row number  (i.e get row number)
	   Row row0 = sheet1.getRow(0);
	   
	   Cell cell00 = row0.getCell(0);	
	   String cellvalue00 = cell00.getStringCellValue();
	   System.out.println("cell00 value is :" + cellvalue00);
	   
	   Cell cell01 = row0.getCell(1);	
	   String cellvalue01 = cell01.getStringCellValue();
	   System.out.println("cell01 value is :" + cellvalue01);
	   
	  /* 
	   // direct
	   String cellvalue10 = sheet1.getRow(1).getCell(0).getStringCellValue();
	   System.out.println("cell10 value is :" + cellvalue10);
*/
   }

}


/*o/p:-- ok
cell00 value is :rameshsoft.selenium
cell01 value is :java1234
cell10 value is :rameshatbtech
*/
