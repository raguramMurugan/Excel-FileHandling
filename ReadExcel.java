package filehandle;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {

	public static void main(String[] args) throws IOException {
		File f=new File("C:\\Users\\HAI\\Training\\java\\datafiles\\Studentdata.xlsx");
		FileInputStream fis=new FileInputStream(f);
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet sheet=(XSSFSheet) workbook.getSheet("Sheet1");
		int rows=sheet.getLastRowNum();
		int cols=sheet.getRow(0).getLastCellNum();
		
		for(int r=0; r<rows; r++)
		{
			XSSFRow row=sheet.getRow(r);
			
			for(int c=0; c<cols; c++)
			{
				XSSFCell column=row.getCell(c);
				switch(column.getCellType())
				{
				case STRING: System.out.print(column.getStringCellValue()); break;
				case NUMERIC: System.out.print(column.getNumericCellValue()); break;
				case BOOLEAN: System.out.print(column.getBooleanCellValue()); break;
				}
				System.out.print(" | ");
			}
			System.out.println();
			
		}
		
	
		
	
		
		
	}

}

