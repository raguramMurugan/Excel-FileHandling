package filehandle;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class PasswordProtected {

	public static void main(String[] args) throws EncryptedDocumentException, IOException {
		File file=new File("C:\\Users\\HAI\\Training\\java\\datafiles\\Studentdata.xlsx");
		FileInputStream fis=new FileInputStream(file);
		String password="Pass@123";
		XSSFWorkbook book=(XSSFWorkbook)WorkbookFactory.create(file, password);
		XSSFSheet sheet=book.getSheet("Sheet1");
		
		Iterator<Row>iterator=sheet.iterator(); //Contains All rows in Sheet
		while(iterator.hasNext())
		{
			XSSFRow row=(XSSFRow) iterator.next(); //Row Value
			Iterator<Cell>cellIterator=row.cellIterator(); //Contains Cells 
			
			while(cellIterator.hasNext())
			{
				XSSFCell cell=(XSSFCell) cellIterator.next(); //Cell Value
				switch(cell.getCellType())
				{
				case STRING:
					System.out.print(cell.getStringCellValue());break;
				case BOOLEAN:
					System.out.print(cell.getBooleanCellValue()); break;
				case NUMERIC:
					System.out.print(cell.getNumericCellValue()); break;
				}
				System.out.print(" || ");
			}
			System.out.println();
		}
		

	}

}
