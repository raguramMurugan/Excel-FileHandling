package filehandle;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelArray {

	public static void main(String[] args) throws IOException {
		
		XSSFWorkbook book=new XSSFWorkbook();
		XSSFSheet sheet=book.createSheet("Employee Details");
		
		ArrayList <Object[]> Designation=new ArrayList<>();
		Designation.add(new Object[] {"Name","Designation","Years Of Experience"});
		Designation.add(new Object[] {"Ram","Associate Engineer",1.5});
		Designation.add(new Object[] {"Sumithra","Engineer-Technology",2});
		Designation.add(new Object[] {"Ramanan","Project Analyst",1});
		Designation.add(new Object[] {"Senthil","Automation Tester",1.3});
		Designation.add(new Object[] {"BalaKumar","Project Lead",3});
		Designation.add(new Object[] {"Saravanan","Senior Business Assosicate",4});
		Designation.add(new Object[] {"Arut Selvan","Quality Analyst",3.5});
		
		int row=0;
		for(Object[] role:Designation)
		{
			XSSFRow rows=sheet.createRow(row++);
			int column=0;
			for(Object Employee:role)
			{
				
				XSSFCell cell=rows.createCell(column++);
				if(Employee instanceof String)
				{
					cell.setCellValue((String)Employee);
				}
				if(Employee instanceof Integer)
				{
					cell.setCellValue((Integer)Employee);
				}
				if(Employee instanceof Double)
				{
					cell.setCellValue((Double)Employee);
				}
				if(Employee instanceof Boolean)
				{
					cell.setCellValue((Boolean)Employee);
				}
			}
		}
		File f=new File("C:\\Users\\HAI\\Training\\java\\datafiles\\Roles.xlsx");
		FileOutputStream fos=new FileOutputStream(f);
		book.write(fos);
		fos.close();
		System.out.println("Roles details has been Updated in Excel Sheet Successfully");
		
		
	}

}
