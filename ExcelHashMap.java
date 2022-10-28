package filehandle;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelHashMap {

	public static void main(String[] args) throws IOException {
		XSSFWorkbook book =new XSSFWorkbook();
		XSSFSheet sheet=book.createSheet();
		
		Map<String, String>map=new HashMap<>();
		map.put("Student Name", "Marks Obtained");
		map.put("Sumithra", "458");
		map.put("Raguram", "477");
		map.put("Ramanan","438");
		map.put("Vishrutha", "438");
		map.put("BhakyaSri", "438");
		
		int rowcount=0;
		for(Entry<String, String> marks:map.entrySet())
		{
			XSSFRow row=sheet.createRow(rowcount++);
			row.createCell(0).setCellValue(marks.getKey());
			row.createCell(1).setCellValue(marks.getValue());
		}
		
		File file=new File("C:\\Users\\HAI\\Training\\java\\datafiles\\Marks.xlsx");
		FileOutputStream fos=new FileOutputStream(file);
		book.write(fos);
		
		book.close();
		fos.close();
		System.out.println("Marks Obtained Successfully");
	}

}
