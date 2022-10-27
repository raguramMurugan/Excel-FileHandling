package filehandle;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteData2 {

	public static void main(String[] args) throws IOException {
		XSSFWorkbook workbook =new XSSFWorkbook();
		XSSFSheet sheet=workbook.createSheet("Sheet2");
		
		Object candidates[][]= {
								{"CandidateName","CompanyName","Package"},
								{"Hareekesh","TCS",3.5},
								{"Kabilan","Agilisys",4.25},
								{"John","Solarisys",2.75},
								{"Niyas Khan","Infosys",3.5},
								{"Manikandan","Hubino",2.75},
								{"Naveen","Hubino",2.75},
								{"Daniel","Capegemini",5.5}
							};
		int rows=candidates.length;
		int cols=candidates[0].length;
		
		
		for(int r=0; r<rows; r++)
		{
			XSSFRow row=sheet.createRow(r);
			for(int c=0; c<cols;c++)
			{
				XSSFCell cell=row.createCell(c);
				Object information=candidates[r][c];
				if(information instanceof String)
				{
					cell.setCellValue((String)information);
				}
				if(information instanceof Boolean)
				{
					cell.setCellValue((Boolean)information);
				}
				if(information instanceof Double)
				{
					cell.setCellValue((Double)information);
				}
			}
		}
		File file=new File("C:\\Users\\HAI\\Training\\java\\datafiles\\Blank.xlsx");
		FileOutputStream fos=new FileOutputStream(file);
		workbook.write(fos);
		fos.close();
		System.out.println("Datas has Successfully Printed");
	}

}
