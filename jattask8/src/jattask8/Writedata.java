package jattask8;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.*;

public class Writedata {

	public static void main(String[] args) throws IOException {
		
		XSSFWorkbook workbook=new XSSFWorkbook();
		XSSFSheet sheet=workbook.createSheet("Emp Info");
		
		Object empdata[][]= {{"Name","Age","Email"},
		{"John Doe", 30 , "john@test.com"},
		{"Jane Doe", 28 , "john@test.com" },
		{"Bob smith" , 35 , "jacky@example.com"},
		{"Swapnil", 37 , "swapnil@example.com"},
		};
		
		
		int rows= empdata.length;
		int cols =empdata [0].length;
		System.out.println(rows);
		System.out.println(cols);
		
		for (int r=0; r<rows;r++) {
			
			XSSFRow row=sheet.createRow(r);
			for (int c=0; c<cols;c++) {
				XSSFCell cell=row.createCell(c);
				
				Object value=empdata[r][c];
				
				if (value instanceof String)
				cell.setCellValue((String)value);
				if (value instanceof Integer)
					cell.setCellValue((Integer)value);
				if (value instanceof Boolean)
					cell.setCellValue((Boolean)value);
				
			}
		}
		String filePath="\\datafile\\writedata.xlsx";
		FileOutputStream outstream=new FileOutputStream(filePath);
		workbook.write(outstream);
		
		outstream.close();
		
		
		System.out.println("employess.xlsx is done succesfully");
		

	}

}
