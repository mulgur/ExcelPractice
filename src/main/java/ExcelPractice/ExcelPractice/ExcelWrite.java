package ExcelPractice.ExcelPractice;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelWrite {
	
	public static void main(String[] args) throws Exception {
		
		File excel = new File("MOCK_DATA.xlsx");
		
		Workbook wb = WorkbookFactory.create(excel);
		
		Sheet she = wb.getSheet("data");
		
		Row row = she.getRow(6);
		
		Cell cell = row.getCell(0);
		
		cell.setCellValue("mustafa");
		

		
		FileOutputStream fos = new FileOutputStream("yeni.xlsx");
		
		cell.setCellValue("merve");
		
		wb.write(fos);
		
		//cell.setCellValue("merve");
		
	
		

		
		
		

		
		
		
	}

}
