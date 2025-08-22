package jattask8;

import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Paths;
import java.io.FileNotFoundException;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Readdata {

	public static void main(String[] args) throws IOException {
		
		String excelFilePath=Paths.get("datafile","test.xlsx").toString();
		FileInputStream inputstream = new FileInputStream(excelFilePath);
		try (XSSFWorkbook workbook = new XSSFWorkbook(inputstream)) {
			//XSSFSheet sheet = workbook.getSheet("Sheet1");
			XSSFSheet sheet = workbook.getSheetAt(0);
			
			int rows=sheet.getLastRowNum();
			int cols=sheet.getRow(1).getLastCellNum();
			System.out.println(rows);
			System.out.println(cols);
			
			for(int r=0; r<=rows;r++) {
				XSSFRow row=sheet.getRow(r) ;
				 for (int c=0; c<cols;c++) {
					 XSSFCell cell=row.getCell(c);	
					 
					 switch (cell.getCellType()) {
					 case NUMERIC : System.out.print(cell.getNumericCellValue());break;
					 case STRING : System.out.print(cell.getStringCellValue());break;
					 case BOOLEAN : System.out.print(cell.getBooleanCellValue());break;
					default:
						break;
					 }
					 System.out.print(" | ");
				 }
				 System.out.println();
			}
		}
		
					

	}

}
