package filehandle;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteData {

	public static void main(String[] args) throws IOException {
		XSSFWorkbook workbook= new XSSFWorkbook();
		XSSFSheet sheet=workbook.createSheet("Student database");
		Object data[] []= { {"Student Name", "Chemistry", "Maths","Physics","ComputerScience"},
							{"Sumithra",88,90,92,95},
							{"Ragu",80,92,88,75},
							{"Senthil",85,94,84,78},
							{"Ragul",81,95,83,80},
							{"Sai",82,95,75,70},
							{"BhakyaSri",94,78,98,90}
				
		                };
		int rows=data.length;
		int cols=data[0].length;
		
		for(int r=0;r<rows;r++)
		{
			XSSFRow row=sheet.createRow(r);
			for(int c=0; c<cols; c++)
			{
				XSSFCell cell=row.createCell(c);
				Object maindata=data[r][c];
				if(maindata instanceof String)
				{
					cell.setCellValue((String)maindata);
				}
				if(maindata instanceof Boolean)
				{
					cell.setCellValue((Boolean)maindata);
				}
				if(maindata instanceof Integer)
				{
					cell.setCellValue((Integer)maindata);
				}
			}
		}
		File file=new File("C:\\Users\\HAI\\Training\\java\\datafiles\\Blank.xlsx");
		FileOutputStream fos=new FileOutputStream(file);
		workbook.write(fos);
		fos.close();
	}

}
