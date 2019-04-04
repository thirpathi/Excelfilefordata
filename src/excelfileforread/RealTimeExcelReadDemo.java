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

public class RealTimeExcelReadDemo {
	
   public static void main(String[] args) throws EncryptedDocumentException, InvalidFormatException, IOException
   {
	   
	FileInputStream fip = new FileInputStream("D:\\rameshsoft\\workspace\\pb8batch1\\src\\com\\rameshsoft\\automation\\testdata\\TestData.xls");
	Workbook workbook = WorkbookFactory.create(fip);
	Sheet sheet1 = workbook.getSheet("ramesh");

	for(int i=0;i<sheet1.getLastRowNum()+1;i++)
	{
		Row row = sheet1.getRow(i);
		
		for(int j=0;j<row.getLastCellNum();j++)
		{
			Cell cell = row.getCell(j);
			if (cell.getCellType()==cell.CELL_TYPE_STRING) 
			{
				String cellValue00 = cell.getStringCellValue();
				System.out.println(cellValue00);
			}
			
			else if (cell.getCellType()==cell.CELL_TYPE_NUMERIC) 
			{
			double cellData = cell.getNumericCellValue();
			System.out.println(cellData);
			}
			
			else if (cell.getCellType()==cell.CELL_TYPE_BOOLEAN) 
			{
			boolean cellData = cell.getBooleanCellValue();	
			System.out.println(cellData);
			}

		 }
	  }
   }
}


/*o/p:-- ok 
rameshsoft.selenium
java1234
rameshatbtech
1234.0
practice
java
practice well
job is nothing
hello
practice more
*/