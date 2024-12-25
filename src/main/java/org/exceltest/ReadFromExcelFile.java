package org.exceltest;

import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadFromExcelFile {

	public static void main(String[] args) throws IOException {

		FileInputStream fis = new FileInputStream(System.getProperty("user.dir") + "\\Exceldata\\sample1.xlsx");
		XSSFWorkbook xsw = new XSSFWorkbook(fis);
		String sheetName = xsw.getSheetName(0);

		System.out.println("First Sheet Name :" + sheetName);

		XSSFSheet sheet = xsw.getSheet(sheetName);
		int totalrows = sheet.getLastRowNum();
	
		int totalcolumns = sheet.getRow(1).getLastCellNum();
		System.out.println("last row : " + totalrows + " , Last column : " + totalcolumns);

		

	for (int r = 0; r <= sheet.getLastRowNum(); r++) {
		if(r==1) {
		Row row = sheet.getRow(1);
//		System.out.println("Row number " + row);
		if(row!=null) {
			for (int c = 1; c <= row.getLastCellNum(); c++) {
				
				if(c==3) {
				// Get the second cell (index 1) in the row
	            Cell cell = row.getCell(c);
				if(cell!=null) {
	            CellType cellType = cell.getCellType();
	            double stringCellValue = cell.getNumericCellValue();
	            System.out.println(cellType + " "+stringCellValue);
	            break;
				}
				}else {
					//System.out.println("column not mached");
				}
			}
			}else {
				System.out.println("Row not available");
				
			}
		}
		
	}

		xsw.close();
		fis.close();

	}

}
