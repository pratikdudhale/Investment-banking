package excelSheetStudy;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Date;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class StudentsDATA {

	public static void main(String[] args) throws EncryptedDocumentException, IOException {
	
		FileInputStream obj=new FileInputStream("C:\\Users\\Admin\\Desktop\\9TH PLATIUM.xlsx");

//		String m = WorkbookFactory.create(obj).getSheet("sheet1").getRow(0).getCell(0).getStringCellValue();
//		
//		System.out.println("Excelsheet value is "+ m);
//		
		Sheet excelsheet = WorkbookFactory.create(obj).getSheet("sheet1");
		
	Date DOB = excelsheet.getRow(2).getCell(2).getDateCellValue();
	
		
	String name = excelsheet.getRow(2).getCell(3).getStringCellValue();
	
	double RollNo = excelsheet.getRow(2).getCell(1).getNumericCellValue();
	
	
	System.out.println("DOB of STUDENT is "+ DOB);
	
	System.out.println("Students name is "+name);
	
	System.out.println("Roll no of  "+ name+ " is " +RollNo);
	}
}

