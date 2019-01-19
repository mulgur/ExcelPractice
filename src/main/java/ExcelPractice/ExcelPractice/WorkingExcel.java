package ExcelPractice.ExcelPractice;

import java.io.File;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class WorkingExcel {

	String filePath = "MOCK_DATA.xlsx";

	public static void main(String[] args) throws Exception {
		
		getAllSheetData("MOCK_DATA.xlsx","data");
System.out.println("***********");
		// Workbook --> Sheet ---> Row--> Cell

		// Earlier version of poi library
		// have 2 different set of classes to work with xls , xlsx files
		/*
		 * xls files --- MS Excel 97-2003 HSSFWorkbook , HSSFSheet , HSSFRow , HSSFCell
		 * xlsx XSSFWorkbook , XSSFSheet , XSSFRow , XSSFCell
		 */

		File excelFile = new File("MOCK_DATA.xlsx");
		Workbook wb = WorkbookFactory.create(excelFile);

		System.out.println(wb.getNumberOfSheets());

		// Sheet sh = wb.getSheet("data");
		Sheet sh = wb.getSheetAt(0);
		Row row1 = sh.getRow(0);
		Cell c1 = row1.getCell(1);
		System.out.println(c1);

		//how many cell in one row
		int columnCountInFirstRow = row1.getLastCellNum();
		System.out.println(columnCountInFirstRow);
		
		
        //how many non-empty row number in one sheet  
		 int rowCount = sh.getLastRowNum();
		 System.out.println( rowCount );

		// getPhysicalNumberOfRows will return actual rowNumber
		// whether you have empty value row or not
		int getNonEmptyRowCount = sh.getPhysicalNumberOfRows();
		System.out.println(getNonEmptyRowCount);

		for (int i = 0; i < getNonEmptyRowCount; i++) {

			Row row = sh.getRow(i);
			
			System.out.println("ROW NUMBER : " + (i+1));


			for (int j = 0; j < columnCountInFirstRow; j++) {

				Cell cell = row.getCell(j);
				System.out.print(cell + "---");

			}
			System.out.println();

		}
		
		


		// Create a utility method to store all sheetData
		// in two dimensional String Array

		// method name : getAllSheetDate
		// return type : String[][]
		// params : FileName as String , SheetName

		getAllSheetData2();
		wb.close();

	}

//	public String[][] getAllSheetDate(String FileName, String SheetName) throws Exception {
//		// File excelFile = new File("MOCK_DATA.xlsx") ;
//		FileInputStream fis = new FileInputStream(filePath);
//		Workbook wb = WorkbookFactory.create(fis);
//		// Sheet sh = wb.getSheet(SheetName);
//		Sheet sheet = wb.getSheet(SheetName);
//		int rowCount = sheet.getPhysicalNumberOfRows();
//		int colCount = sheet.getRow(0).getLastCellNum();
//
//		String[][] data = new String[rowCount][colCount];
//
//		for (int i = 0; i < rowCount; i++) {
//			System.out.println("row number : " + i);
//
//			for (int j = 0; j < colCount; j++) {
//				Cell cell = sheet.getRow(i).getCell(j);
//				// System.out.println( cell.toString());
//				data[i][j] = cell.toString();
//			}
//			// System.out.println();
//		}
//		fis.close();
//		wb.close();
//		return data;
//	}
	
	public static String[][] getAllSheetData(String FileName, String SheetName) throws Exception{
		File excelFile = new File(FileName);
		
		Workbook wb = WorkbookFactory.create(excelFile);
		
		Sheet sheet = wb.getSheet(SheetName);
		
		
		int rowCount = sheet.getPhysicalNumberOfRows();
		int columnCount = sheet.getRow(0).getLastCellNum();
		
		String[][] data = new String[rowCount][columnCount];
		
		for(int i = 0 ; i < rowCount; i++) {
			System.out.println("row num : " + (i+1));
			
			for(int j = 0; j < columnCount; j++) {
				Cell cell = sheet.getRow(i).getCell(j);
				System.out.print(cell+ "-----");
				data[i][j] = cell.toString();
				
			}
			System.out.println( );
		}
		
		wb.close();
		return data;
		
	}
	
	

	public static void getAllSheetData2() throws Exception {
		File excelFile = new File("MOCK_DATA.xlsx");
		Workbook wb = WorkbookFactory.create(excelFile);

		Sheet sheet = wb.getSheetAt(0);
		int rowCount = sheet.getPhysicalNumberOfRows();
		int colCount = sheet.getRow(0).getLastCellNum();

		for (int i = 0; i < rowCount; i++) {
			System.out.println("row number : " + i);

			for (int j = 0; j < colCount; j++) {
				Cell cell = sheet.getRow(i).getCell(j);
				System.out.println(cell.toString());
			}
			System.out.println();

		}
		wb.close();
	}
}