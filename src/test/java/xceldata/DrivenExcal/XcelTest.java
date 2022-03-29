package xceldata.DrivenExcal;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XcelTest {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

		
		FileInputStream fis=new FileInputStream("D:\\Book11.xlsx");
		XSSFWorkbook  book=new XSSFWorkbook(fis);
		
		XSSFSheet sheet=book.getSheetAt(0);
		
	/*	int row=sheet.getLastRowNum();
		int col=sheet.getRow(0).getLastCellNum();
		
		for(int r=0;r<row;r++) {
			
			
		XSSFRow	 rows=sheet.getRow(r);
			
			for(int c=0;c<col;c++) {
				
				XSSFCell cells=rows.getCell(c);
				
				
				switch(cells.getCellType()){
				case STRING: System.out.println(cells.getStringCellValue());break;
				case NUMERIC: System.out.println(cells.getNumericCellValue());break;
				case BOOLEAN: System.out.println(cells.getBooleanCellValue());break;
				}
				
			}
			System.out.println();
		}*/
		
		// using iterator
		Iterator iterate=sheet.rowIterator();
		while(iterate.hasNext()) {
			
			XSSFRow row=(XSSFRow) iterate.next();
			
			Iterator itr=row.cellIterator();
			while(itr.hasNext()) {
				 XSSFCell cells=(XSSFCell) itr.next();
				 
				 switch(cells.getCellType()){
					case STRING: System.out.print(cells.getStringCellValue());break;
					case NUMERIC: System.out.print(cells.getNumericCellValue());break;
					case BOOLEAN: System.out.print(cells.getBooleanCellValue());break;
					}
					System.out.print("  /  ");
				}
				System.out.println();

				 
			}
			
		}
		
		
	}
