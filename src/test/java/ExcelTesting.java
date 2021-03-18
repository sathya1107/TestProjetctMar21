import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Date;
import java.text.SimpleDateFormat;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.SystemOutLogger;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelTesting {

	

	public static void main(String[] args) throws IOException {
		File f = new File("C:\\Users\\sathya\\eclipse-workspace\\DataDrive\\Excel\\TestExcel.xlsx");
		FileInputStream excelLoc = new FileInputStream(f);
		Workbook w = new XSSFWorkbook(excelLoc);
		Sheet s = w.getSheet("sheet1");
		for (int i = 0; i < s.getPhysicalNumberOfRows(); i++) {
			Row r = s.getRow(i);
			for (int j = 0; j <r.getPhysicalNumberOfCells() ; j++) {
				Cell c = r.getCell(j);
				int ct = c.getCellType();
				if (ct==1) {
					String stri = c.getStringCellValue();
					System.out.println(stri);
				}
				if (ct==0) {
					if (DateUtil.isCellDateFormatted(c)) {
						    Date date = c.getDateCellValue();
						   SimpleDateFormat ss = new SimpleDateFormat("dd-MM-yyyy");
						    System.out.println(ss.format(date));
								 
						
					} else {
					long l =(long)c.getNumericCellValue();
					System.out.println(l);
					

					}
					
					
				}
					
								
				
				
				
			}
		}
		
		
		
		
		
		

	}

	private static String String(double numericCellValue) {
		// TODO Auto-generated method stub
		return null;
	}

}
