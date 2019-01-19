package ExcelPractice.ExcelPractice;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Arrays;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelStuff {
	static String filePath ="MOCK_DATA.xlsx";

  public static void main(String[] args) throws Throwable {
    
//    printAllSheetData();
//    
//    String[][] result = getAllSheetDate("MOCK_DATA.xlsx","data");
//    
//    
//    		System.out.println(Arrays.deepToString(result));
    		
    		System.out.println(getCellData("MOCK_DATA.xlsx","data",5,1));

  }
  
  public static void printAllSheetData() throws Exception {
    
    File excelFile = new File("MOCK_DATA.xlsx") ; 
    Workbook wb = WorkbookFactory.create(excelFile); 
    
    Sheet sheet = wb.getSheetAt(0); 
    int rowCount = sheet.getPhysicalNumberOfRows();
    int colCount = sheet.getRow(0).getLastCellNum();
    
    for (int i = 0; i < rowCount; i++) {
      
      System.out.println(" row number : "+ i);
      
      for (int j = 0; j < colCount; j++) {
        
        Cell cell = sheet.getRow(i).getCell(j); 
        System.out.print( cell.toString() + " | ");
        
      }
      System.out.println();
      
      
    }

  }
  
  public static String[][]  getAllSheetData(String FileName, String SheetName) throws Exception{
		 //File excelFile = new File("MOCK_DATA.xlsx") ; 
		 FileInputStream fis = new FileInputStream (filePath); 
		 Workbook wb = WorkbookFactory.create(fis);
		// Sheet sh = wb.getSheet(SheetName); 
		 Sheet sheet = wb.getSheet(SheetName);
		 int rowCount = sheet.getPhysicalNumberOfRows();
		 int colCount = sheet.getRow(0).getLastCellNum();
		 
		 String[][] data = new String [rowCount] [colCount];
		 
		 for(int i=0; i<rowCount; i++) {
			 System.out.println("row number : "+i);
			 
			 for(int j=0; j<colCount; j++) {
				 Cell cell=sheet.getRow(i).getCell(j);
				// System.out.println( cell.toString());
				 data[i][j] = cell.toString();
			 }
		// System.out.println();
	}
		 fis.close();
		 wb.close();
		 return data;
	 }

  public static String getCellData(String filePath, String sheetname, int rowIndex, int colIndex) throws Throwable {
//	  String result="";
//	  File excelFile = new File("MOCK_DATA.xlsx") ; 
//	    Workbook wb = WorkbookFactory.create(excelFile); 
//	    
//	    Sheet sheet = wb.getSheetAt(0); 
//	    int rowCount = sheet.getPhysicalNumberOfRows();
//	    int colCount = sheet.getRow(0).getLastCellNum();
//	    
//	    for (int i = 0; i < rowCount; i++) {	      
//	      for (int j = 0; j < colCount; j++) {
//	        
//	        Cell cell = sheet.getRow(i).getCell(j); 
//	        result =cell.toString() + " | ";
//	        
//	      }
//	    }
//	  
//	  return result;	 
         String [][] result = getAllSheetData(filePath, sheetname);
	  		  return result[rowIndex][colIndex] ; 
	  


  }

}