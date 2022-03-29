package xceldata.DrivenExcal;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteexcelTest {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

		
		
		XSSFWorkbook workbook=new XSSFWorkbook();
		XSSFSheet sheet=workbook.createSheet("writing");
		
		/*Object empdata[][]= { {"Empid","Empname","jpb"},
				              {1,"naveen","architect"},
				              {1,"harish","tester"},
				              {1,"shankar","developer"},
				            };*/
		
		
		
		ArrayList<Object[]> empdata=new ArrayList<Object[]>();
		
		empdata.add(new Object[] {"Empid","Empname","jpb"});
		empdata.add(new Object[] {1,"naveen","architect"});
		empdata.add(new Object[] {2,"hasii","tester"});
		empdata.add(new Object[] {1,"shankar","developer"});
		
		/*int rows=empdata.length;
		int cols=empdata[0].length;
		
		
		for(int r=0;r<rows;r++) {
			
		XSSFRow	row=sheet.createRow(r);
			
			for(int c=0;c<cols;c++) {
			XSSFCell cell=row.createCell(c);
			
			 Object value=empdata[r][c];
			 
			if(value instanceof String) 
				cell.setCellValue((String)value);
				
			if(value instanceof Integer) 
					cell.setCellValue((Integer)value);
			if(value instanceof Boolean) 
				cell.setCellValue((Boolean)value);		
				
			
	}
			
		}*/
		
		// using for each loop
		
		int rowCount=0;
		for(Object[] empda:empdata) {
			
			XSSFRow rows=sheet.createRow(rowCount++);
			
			
			int cellCount=0;
			for(Object value:empda) {
				
				 XSSFCell cell=rows.createCell(cellCount++);
				 
				 if(value instanceof String) 
						cell.setCellValue((String)value);
						
					if(value instanceof Integer) 
							cell.setCellValue((Integer)value);
					if(value instanceof Boolean) 
						cell.setCellValue((Boolean)value);
			}
		}
		
		FileOutputStream fos=new FileOutputStream("D:\\writeexcel.xlsx");
		workbook.write(fos);
		
		fos.close();
		System.out.println("we have successfully written data in excel sheet");
	}
			
		

}
