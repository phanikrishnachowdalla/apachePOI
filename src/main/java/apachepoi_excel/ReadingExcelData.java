package apachepoi_excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.formula.atp.Switch;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadingExcelData {

	/**
	 * @param args
	 * @throws IOException
	 */
	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		String excelFilePath = "C:\\Users\\DELL\\eclipse-workspace\\ApachePOI\\datafiles\\my_friends.xlsx";
		
		FileInputStream fip = new FileInputStream(excelFilePath);
		XSSFWorkbook workbook = new XSSFWorkbook(fip);
		XSSFSheet sheet = workbook.getSheet("Sheet1");
		
//		System.out.println(sheet = workbook.getSheetAt(0));
//		System.out.println(workbook.getSheetIndex("Sheet1"));
//		System.out.println(workbook.getSheetName(0));
		
	/*	int rows = sheet.getLastRowNum();
		System.out.println(rows);
		int cols = sheet.getRow(1).getLastCellNum();
		System.out.println(cols);
		
		// using FOR loop
		for(int r=0;r<=rows;r++)
		{
			XSSFRow row = sheet.getRow(r);
			for(int c=0;c<cols;c++)
			{
				XSSFCell cell = row.getCell(c);
				
				switch(cell.getCellType())
				{
				case STRING: System.out.print(cell.getStringCellValue()); break;
				case NUMERIC: System.out.print(cell.getNumericCellValue()); break;
				case BOOLEAN: System.out.print(cell.getBooleanCellValue()); break;
				}
				System.out.print(" | ");
			}
			System.out.println();
		} */
		
		//using ITERATOR
		
		Iterator iterator = sheet.iterator();
		
		while(iterator.hasNext())
		{
			XSSFRow row = (XSSFRow) iterator.next();
			Iterator cellIterator = row.cellIterator();
			while(cellIterator.hasNext())
			{
				XSSFCell cell = (XSSFCell) cellIterator.next();
				switch(cell.getCellType())
				{
				case NUMERIC: System.out.print(cell.getNumericCellValue()); break;
				case STRING: System.out.print(cell.getStringCellValue()); break;
				case BOOLEAN: System.out.print(cell.getBooleanCellValue()); break;
				}
				System.out.print(" | ");
			}
			System.out.println();
		}
		
	}

}
