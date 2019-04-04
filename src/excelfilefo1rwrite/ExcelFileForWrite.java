package excelfilefo1rwrite;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelFileForWrite {
	
  public static void main(String[] args) throws IOException, EncryptedDocumentException, InvalidFormatException {

	FileInputStream fip = new FileInputStream("D:\\rameshsoft\\workspace\\pb8batch1\\src\\com\\rameshsoft\\automation\\testdata\\TestData.xls");
    Workbook workbook = WorkbookFactory.create(fip);
    
    // create sheet name (i.e getsheet)
    Sheet sheet1 = workbook.getSheet("tejas");
    
    // create row number  (i.e select row number)
    Row row3 = sheet1.createRow(3);
	
    // select cell (i.e give the cell number)
	Cell cell30 = row3.createCell(0);
	cell30.setCellValue("practice well");
	
	Cell cell31 = row3.createCell(1);
	cell31.setCellValue("job is nothing");
	
	/*
    // direct one step 
    // first select row on that give cell number on that pass the message	
	value = sheet1.createRow(4)createCell(0).setCellValue("hello");
	*/
	
	//store the excel file
	FileOutputStream fop = new FileOutputStream("D:\\rameshsoft\\workspace\\pb8batch1\\src\\com\\rameshsoft\\automation\\testdata\\TestData.xls");
	workbook.write(fop);
	
  }

}


/*o/p:--ok
okkkk*/

