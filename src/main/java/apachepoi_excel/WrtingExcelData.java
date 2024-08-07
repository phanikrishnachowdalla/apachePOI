package apachepoi_excel;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

//workbook->sheet->row->column

public class WrtingExcelData {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Sheet1");
		
		
		//////////////////////////////////////////////////////////////////////
		// ARRAY-LIST
		ArrayList<Object[]> empdata = new ArrayList<Object[]>();
		empdata.add(new Object[] {"empid","name","life-partner"});
		empdata.add(new Object[] {14,"honey","phani"});
		empdata.add(new Object[] {42,"phani","honey"});
		empdata.add(new Object[] {1,"komali","prabhas"});
		empdata.add(new Object[] {2,"prabhas","komali"});
		
		// using FOR-EACH loop
		int rowNum=0;
		for(Object[] emp:empdata)
		{
			XSSFRow row = sheet.createRow(rowNum++);
			int cellNum=0;
			for(Object value:emp)
			{
				
				XSSFCell cell = row.createCell(cellNum++);
				if(value instanceof String)
					cell.setCellValue((String) value);
				if(value instanceof Boolean)
					cell.setCellValue((boolean) value);
				if(value instanceof Integer)
					cell.setCellValue((Integer) value);
				
				
			}
		}
		
		
		//////////////////////////////////////////////////////////////////////
		
		// 2-dimensional ARRAY
//		Object Empdata[][] = {  {"EmpID","Name","Gender","life-Partner"},
//								{ 1,"Honey","Femail","Phani"},
//								{2,"Phani","Mail","Honey"},
//								{3,"Komali","Femail","Some X"}
//							 };
		
		//using normal FOR loop
	/*	int rows = Empdata.length;
		int cols = Empdata[0].length;
		System.out.println(rows);
		System.out.println(cols);
		
		for(int r=0;r<rows;r++)
		{
			XSSFRow newRow = sheet.createRow(r);
			
			for(int c=0;c<cols;c++)
			{
				XSSFCell newCell = newRow.createCell(c);
				Object value = Empdata[r][c];
				
				if(value instanceof String)
					newCell.setCellValue((String)value);
				if(value instanceof Boolean)
					newCell.setCellValue((Boolean)value);
				if(value instanceof Integer)
					newCell.setCellValue((Integer)value);
			}
		} */
		
		// using FOR-EACH loop
//		int rowCount = 0;
//		for(Object emp[]:Empdata)
//		{
//			XSSFRow row = sheet.createRow(rowCount++);
//			int columnCount = 0;
//			
//			for(Object value:emp)
//			{
//				XSSFCell cell = row.createCell(columnCount++);
//				
//				if(value instanceof String)
//					cell.setCellValue((String)value);
//				if(value instanceof Integer)
//					cell.setCellValue((Integer)value);
//				if(value instanceof Boolean)
//					cell.setCellValue((Boolean)value);
//			}
//		}
		
		String filePath = "C:\\Users\\DELL\\eclipse-workspace\\Apache_POI\\datafiles\\life_partners.xlsx";
		FileOutputStream fop = new FileOutputStream(filePath);
		workbook.write(fop);
		
		fop.close();
		System.out.println("Grand Success");
		
	}

}
