package excelSheetStudy;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.streaming.DeferredSXSSFWorkbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

public class NewExcel {

	public static void main(String[] args) throws EncryptedDocumentException, IOException {
	
	
	
	
	//create an Object of file input stream and give path along with file name and extension
			FileInputStream Myfile= new FileInputStream("C:\\Users\\Admin\\Documents\\School Uniform 2022-23.xlsx");

		//	String value = WorkbookFactory.create(Myfile).getSheet("Sheet1").getRow(2).getCell(0).getStringCellValue();
			
			//System.out.println("Data from excel is "+value);
			
		//	  double value2 = WorkbookFactory.create(Myfile).getSheet("Sheet1").getRow(4).getCell(0).getNumericCellValue();
			//System.out.println("Data from excel is "+value2);
			
			// WorkbookFactory--> will return workbook 
		Workbook test = WorkbookFactory.create(Myfile);
		//get sheet will return sheet type 
			Sheet MySheet = test.getSheet("Sheet1");
//			//get row will return a row type
			 Row myRow = MySheet.getRow(1);
			//get cell will return cell type
			 Cell Mycell = myRow.getCell(1);
			//getStringCellValue will return String type value
			String MyValue = Mycell.getStringCellValue();
			System.out.println(MyValue);
			
			
		}


	

}

