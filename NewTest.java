package keyword.keyword;


import java.io.File;
import java.io.IOException;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.testng.annotations.Test;

import keyword.Comparefile;

public class NewTest {
	
		 
		@Test
		public void differentNumberOfSheets() throws EncryptedDocumentException, IOException
		{
			String userDir = System.getProperty("user.dir");
			Workbook wb1 = WorkbookFactory.create(new File(userDir+"\\src\\Data\\ExcelFilesWithThreeSheets.xlsx"));
			Workbook wb2 = WorkbookFactory.create(new File(userDir+"\\src\\Data\\ExcelFilesWithTwoSheets-Data-Actual"));
			Comparefile Comparefile = new Comparefile();
			Comparefile.verifyIfExcelFilesHaveSameNumberAndNameOfSheets(wb1, wb2);
			
		}
		
		@Test
		public void sameNumberOfSheets() throws EncryptedDocumentException, IOException
		{
			String userDir = System.getProperty("user.dir");
			Workbook wb1 = WorkbookFactory.create(new File(userDir+"\\src\\Data\\ExcelFilesWithTwoSheets-Data-Actual.xlsx"));
			Workbook wb2 = WorkbookFactory.create(new File(userDir+"\\src\\Data\\ExcelFilesWithTwoSheets-Data-Match.xlsx"));
			Comparefile Comparefile = new Comparefile();
			Comparefile.verifyIfExcelFilesHaveSameNumberAndNameOfSheets(wb1, wb2);
		}
		
		
		@Test
		public void differentNumberOfRows() throws EncryptedDocumentException, IOException
		{
			String userDir = System.getProperty("user.dir");
			Workbook wb1 = WorkbookFactory.create(new File(userDir+"\\src\\Data\\ExcelFilesWithTwoSheets-Data-Actual.xlsx"));
			Workbook wb2 = WorkbookFactory.create(new File(userDir+"\\src\\Data\\ExcelFilesWithTwoSheets-Extra-Row.xlsx"));
			Comparefile Comparefile = new Comparefile();
			Comparefile.verifySheetsInExcelFilesHaveSameRowsAndColumns(wb1, wb2);
		}
		
		@Test
		public void differentNumberOfColumns() throws EncryptedDocumentException, IOException
		{
			String userDir = System.getProperty("user.dir");
			Workbook wb1 = WorkbookFactory.create(new File(userDir+"\\src\\Data\\ExcelFilesWithTwoSheets-Data-Actual.xlsx"));
			Workbook wb2 = WorkbookFactory.create(new File(userDir+"\\src\\Data\\ExcelFilesWithTwoSheets-Extra-Columns.xlsx"));
			Comparefile Comparefile = new Comparefile();
			Comparefile.verifySheetsInExcelFilesHaveSameRowsAndColumns(wb1, wb2);
		}
		
		@Test
		public void sameContentOfExcelFiles() throws EncryptedDocumentException, IOException {
			String userDir = System.getProperty("user.dir");
			Workbook wb1 = WorkbookFactory.create(new File(userDir+"\\src\\Data\\ExcelFilesWithTwoSheets-Data-Actual.xlsx"));
			Workbook wb2 = WorkbookFactory.create(new File(userDir+"\\src\\Data\\ExcelFilesWithTwoSheets-Data-Mismatch.xlsx"));
			Comparefile Comparefile = new Comparefile();
			Comparefile.verifyIfExcelFilesHaveSameNumberAndNameOfSheets(wb1, wb2);
			Comparefile.verifySheetsInExcelFilesHaveSameRowsAndColumns(wb1, wb2);
			Comparefile.verifyDataInExcelBookAllSheets(wb1, wb2);
		
		}
		
		
	}
