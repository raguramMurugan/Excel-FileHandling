package filehandle;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Pricing {

	public static void main(String[] args) throws IOException {
		
		File file =new File("C:\\Users\\HAI\\Training\\java\\datafiles\\Product.xlsx");
		FileInputStream fis=new FileInputStream(file);
		XSSFWorkbook book=new XSSFWorkbook(fis);
		XSSFSheet sheet=book.getSheet("Sheet1");
		
		
		sheet.getRow(8).getCell(1).setCellFormula("SUM(B1:B8)");
		fis.close();
		
		FileOutputStream fos=new FileOutputStream(file);
		book.write(fos);
		book.close();
		fos.close();
		
		System.out.println("Bill has been Updated");
	}

}