package excelSheetStudy;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class AdmissionForm {

	public static void main(String[] args) throws EncryptedDocumentException, IOException {
		
		FileInputStream obj = new FileInputStream("C:\\Users\\Admin\\Documents\\Admission Form Checklist 2021.xlsx");

//		String name = WorkbookFactory.create(obj).getSheet("sheet1").getRow(1).getCell(0).getStringCellValue();
//		System.out.println(name);

		
		Sheet excelsheet = WorkbookFactory.create(obj).getSheet("sheet2");
		
		for(int i=0;i<=8;i++) {
			String value = excelsheet.getRow(i).getCell(0).getStringCellValue();
			System.out.println(value);
		}
		for(int i=0;i<=8;i++) {
			String value = excelsheet.getRow(0).getCell(i).getStringCellValue();
			System.out.println(value);
		}
		
		
		// To get Last row no
		
		int lastcellno = excelsheet.getLastRowNum();
		
		
	System.out.println(	lastcellno);
	int total = lastcellno-1;
	
	for(int i=0;i<=total;i++) {
		String v = excelsheet.getRow(1).getCell(i).getStringCellValue();
		System.out.println(v);
		
		
	}
	
		
		
	
		
		
	
	
	}

}
