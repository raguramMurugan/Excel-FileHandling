package filehandle;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class MappingData {

	public static void main(String[] args) throws InvalidFormatException, IOException {
		
		
		XSSFWorkbook book=new XSSFWorkbook();
		XSSFSheet sheet=book.createSheet("Sheet1");
		
		Map<String, String>map=new HashMap<>();
		map.put("ASUS Tuf FX504","75000");
		map.put("ASUS Rog", "85000");
		map.put("Predator", "95000");
		map.put("MacBook Pro", "80000");
		map.put("Aser", "55000");
		map.put("HP", "35000");
		map.put("Dell", "75000");
		map.put("Dell Latitude", "25000");
		
		int rowcount=0;
		for(Entry<String, String> product:map.entrySet())
		{
			XSSFRow row=sheet.createRow(rowcount++);
			row.createCell(0).setCellValue(product.getKey());
			row.createCell(1).setCellValue(product.getValue());			
		}
		
		
		File file=new File("C:\\Users\\HAI\\Training\\java\\datafiles\\Product.xlsx");
		FileOutputStream fos=new FileOutputStream(file);
		book.write(fos);
		book.close();
		fos.close();
		System.out.println("Successfully done");
		
	}

}
