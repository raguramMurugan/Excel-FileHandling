package filehandle;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class GetData {

	public static void main(String[] args) throws IOException {
		File file=new File("C:\\Users\\HAI\\Training\\java\\datafiles\\Marks.xlsx");
		FileInputStream fis=new FileInputStream(file);
		XSSFWorkbook book =new XSSFWorkbook(fis);
		XSSFSheet sheet=book.getSheet("Sheet0");
		HashMap<String, String>map=new HashMap<>();
		
		int rowcount=sheet.getLastRowNum();
		
		for(int r=0;r<rowcount; r++)
		{
			String key=sheet.getRow(r).getCell(0).getStringCellValue();
			String value=sheet.getRow(r).getCell(1).getStringCellValue();
			map.put(key, value);	
		}
		for(Entry datas:map.entrySet())
		{
			System.out.print(datas.getKey()+"  ||  ");
			System.out.print(datas.getValue());
			System.out.println();
		}
	
		book.close();
		fis.close();
	}

}
