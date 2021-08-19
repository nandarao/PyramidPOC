package utility;



import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.Calendar;
import java.util.Hashtable;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ReadAndWrite_xlsx {

public String FileLoation;
FileInputStream FIS;
FileOutputStream FOS;
XSSFWorkbook WB;
XSSFSheet WSName;



/**
 * @param FileLoation file path of excel sheet
 * 
 */
public ReadAndWrite_xlsx(String FileLoation) {
	this.FileLoation=FileLoation;
	try{
		FIS = new FileInputStream(new File(FileLoation));
			WB=new XSSFWorkbook(FIS);
		
		WSName = WB.getSheetAt(0);
		if(WSName == null){
			FIS.close();
			System.out.println("No Sheet With This name");
		}
		FIS.close();
		}
	catch(FileNotFoundException e){
		e.printStackTrace();
	}catch (IOException e) {
		e.printStackTrace();
	}catch (Exception e){
			e.printStackTrace();
		}
		//return path;
	
}



/**
 * this method reads Test data values as per test case
 * @param SheetName SheetName -- Test Data
 * @param HeaderOrTastCaseName Header and Test case name are same in  Test Data sheet   
 * @return returns values with Column name and its values (i.e url = wwww.xxxx.com)
 */
public  Object[][]  retrieveTestDataValues( String SheetName,String HeaderOrTastCaseName){
	int ColHeaderNo=0;
	int HeaderRowNo = 0;
	int count=0;
	int countNo=-1;
	int TotalColNO = 0;
	Object[][] data=null;
	Hashtable<String, String>table=null;
	
	WSName = WB.getSheet(SheetName);
	
	int TotalNoRows = WSName.getLastRowNum()+1;
//	System.out.println("Total Rows"+TotalNoRows);
	if (TotalNoRows==0) {
//				System.out.println("null");
				return null;
	}
	
	for (int i = 0; i < TotalNoRows; i++) {
		XSSFRow Row = WSName.getRow(i);
		if (Row==null) {
			continue;
		}
		XSSFCell cell = Row.getCell(ColHeaderNo);
//		System.out.println("Header Name "+cell.getStringCellValue());
		if (cell.getStringCellValue().trim().equalsIgnoreCase(HeaderOrTastCaseName)) {
			HeaderRowNo=i+1;
			TotalColNO = WSName.getRow(HeaderRowNo).getLastCellNum();
//			System.out.println("header row no "+HeaderRowNo);
			XSSFRow Row1 = WSName.getRow(HeaderRowNo);
			XSSFCell cell2 = Row1.getCell(0);
//			System.out.println("col no"+TotalColNO+cell2);
			break;
		}
	}
close:
	for (int j = (HeaderRowNo+1); j < TotalNoRows; j++) {
		XSSFRow Row =WSName.getRow(j);
		XSSFCell cell = Row.getCell(ColHeaderNo);
		//System.out.println("header no "+(HeaderRowNo+1));
//		System.out.println("run mode"+cell.getStringCellValue());
		if (cell.getStringCellValue().trim().equalsIgnoreCase("y") || cell.getStringCellValue().trim().equalsIgnoreCase("n")) {
			count++;
		}
		else { cell.getStringCellValue().trim().equalsIgnoreCase("");
			break close;
	}

	}
//	System.out.println("count "+count);
	
	data=new Object[count][1];
	
	
	for (int k =(HeaderRowNo+1) ;k < (HeaderRowNo+count+1); k++) {
		table= new Hashtable<String,String>();
		countNo++;
	for (int L = 0; L < TotalColNO; L++) {
		table.put(cellText((HeaderRowNo),L), cellText(k, L));
	}	
	data[countNo][0]=table;
	/*			XSSFRow Row =WSName.getRow(k);
		for (int m = 0; m < TotalColNO; m++) {
			XSSFCell cell = Row.getCell(m);
			System.out.println(cell.getStringCellValue());
			
		}*/
	}
	return data;
}


/**
 * this method used for read values of excel data types like string, number, date
 * @param rowNo Row Number of excel
 * @param colNo column number of excel 
 * @return returns any value as string 
 */
public  String cellText(int rowNo, int colNo){
	String cellText = null;
	
	XSSFRow row = WSName.getRow(rowNo);
	
	if (row==null) {
//	System.out.println("file closed");
//	System.out.println("no of row"+row);
	
	return null;
	}
	XSSFCell cell=row.getCell(colNo);
	if (cell==null) {
		return null;
	}
	
	//System.out.println(cell.getStringCellValue());
	
	switch (cell.getCellType()) {
	case Cell.CELL_TYPE_BLANK:
		cellText = String.valueOf("");
	//	System.out.println(cellText);
		break;
	case Cell.CELL_TYPE_BOOLEAN:
		cellText = String.valueOf(cell.getBooleanCellValue());
		//System.out.println(cellText);
		break;
	case Cell.CELL_TYPE_ERROR:
		cellText = String.valueOf(cell.getErrorCellValue());
		//System.out.println(cellText);
		break;
	case Cell.CELL_TYPE_FORMULA:
		cellText = String.valueOf(cell.getNumericCellValue());
		//System.out.println(cellText);
		break;
	case Cell.CELL_TYPE_NUMERIC:
	
		if (HSSFDateUtil.isCellDateFormatted(cell)) {
			double d =cell.getNumericCellValue();
			Calendar cal =Calendar.getInstance();
			cal.setTime(HSSFDateUtil.getJavaDate(d));
			cellText = cal.get(Calendar.DAY_OF_MONTH)+"-"+(cal.get(Calendar.MONTH)+1)+"-"+cal.get(Calendar.YEAR);
			//System.out.println("date value " +cellText);
			
		} else {
			BigDecimal number = new BigDecimal(cell.getNumericCellValue());	
			//int number = new BigDecimal(cell.getNumericCellValue()).intValue();		
			cellText = String.valueOf(number);
			//System.out.println("number value "+cellText);
		}
		
		
		break;
	case Cell.CELL_TYPE_STRING:
		cellText = String.valueOf(cell.getStringCellValue());
		//System.out.println(cellText);
		break;
	}
	return cellText;
	
	//return cellText;
}

/**
 * this method is used for read runmode of test data 
 * Access with Constructor ReadAndWrite with file location 
 * @param SheetName Sheet Name -- Test Data
 * @param HeaderOrTastCaseName  Header and Test case name are same in  Test Data sheet   
 * @param colName Runmode 
 * @return  array stores values of Test case runmode
 */
public String[] retrieveToRunFlagTestData(String SheetName, String HeaderOrTastCaseName, String colName){
	int ColHeaderNo=0;
	int HeaderRowNo = 0;
	int count=0;
	int countNo=-1;
	int TotalColNO = 0;
	int insertValueInToArray=-1;
	WSName=WB.getSheet(SheetName);
	int TotalNoRows = WSName.getLastRowNum()+1;
//	System.out.println("Total Rows"+TotalNoRows);
	if (TotalNoRows==0) {
		//return null;
//		System.out.println("null");
	}
	
	for (int i = 0; i < TotalNoRows; i++) {
		XSSFRow Row = WSName.getRow(i);
		if (Row==null) {
			continue;
		}
		XSSFCell cell = Row.getCell(ColHeaderNo);
//		System.out.println("Header Name "+cell.getStringCellValue());
		if (cell.getStringCellValue().trim().equalsIgnoreCase(HeaderOrTastCaseName)) {
			HeaderRowNo = i+1;
			TotalColNO = WSName.getRow(HeaderRowNo).getLastCellNum();
//			System.out.println("header row no "+HeaderRowNo);
			XSSFRow Row1 = WSName.getRow(HeaderRowNo);
			XSSFCell cell2 = Row1.getCell(0);
//			System.out.println("col no"+TotalColNO+cell2);
			break;
		}
	}
	close:
		for (int j = (HeaderRowNo+1); j < TotalNoRows; j++) {
			XSSFRow Row =WSName.getRow(j);
			XSSFCell cell = Row.getCell(ColHeaderNo);
			//System.out.println("header no "+(HeaderRowNo+1));
//			System.out.println("run mode"+cell.getStringCellValue());
			if (cell.getStringCellValue().trim().equalsIgnoreCase("y") || cell.getStringCellValue().trim().equalsIgnoreCase("n")) {
				count++;
			}
			else { cell.getStringCellValue().trim().equalsIgnoreCase("");
				break close;
		}

	}
//	System.out.println("count No "+count);
	String data[] = new String[count];
	
	XSSFRow HeaderRow = WSName.getRow(HeaderRowNo);
	for (int i = 0; i < TotalColNO; i++) {
	XSSFCell cell=	HeaderRow.getCell(i);
	
	if (cell.getStringCellValue().trim().equalsIgnoreCase(colName)) {
//		System.out.println("cell Value"+cell.getStringCellValue());
		countNo=i;
	}
	}
	
	if (countNo==-1) {
//		System.out.println("Count No Nullllllllllllll");
		//return null;
	}
	
	for (int j = (HeaderRowNo+1); j <(HeaderRowNo+count+1)  ; j++) {
		insertValueInToArray++;
		XSSFRow Row = WSName.getRow(j);
		if(Row==null){
			data[insertValueInToArray] = "";
		}
		else{
			XSSFCell cell = Row.getCell(countNo);
			if(cell==null){
				data[insertValueInToArray] = "";
			}
			else{
				String value = cell.getStringCellValue();
//				System.out.println("value "+value);
				data[insertValueInToArray] = value;	
			}	
		}
	
		
		/*XSSFCell cell =Row.getCell(countNo);
		System.out.println("Run Mode "+cell.getStringCellValue());*/
	}
	return data;
	
	}


/**
 * This method is used for runmode of Test Case List
 * @param SheetName Sheet Name -- Test Case List
 * @param Header Header 
 * @param TestCaseName Test Case Name 
 * @param colName Runmode 
 * @return retrun value as per Header under Test Case name Runmode
 */
public  String retrieveToRunFlagTestCase(String SheetName, String Header, String TestCaseName, String colName){
	int ColCheckMode=0;
	int TestCaseRowNo = 0;
	int count=0;
	int countNo=-1;
	int TotalColNO = 0;
	WSName=WB.getSheet(SheetName);
	int TotalNoRows = WSName.getLastRowNum()+1;
//	System.out.println("Total Rows"+TotalNoRows);
	if (TotalNoRows==0) {
		return null;
	}
	
	for (int i = 0; i < TotalNoRows; i++) {
		XSSFRow Row = WSName.getRow(i);
		if (Row==null) {
			continue;
		}
		XSSFCell cell = Row.getCell(ColCheckMode);
//		System.out.println("Header Name "+cell.getStringCellValue());
		if (cell.getStringCellValue().trim().equalsIgnoreCase(Header)) {
			TestCaseRowNo = i+1;
			TotalColNO = WSName.getRow(TestCaseRowNo).getLastCellNum();
//			System.out.println("header row no "+TestCaseRowNo);
			XSSFRow Row1 = WSName.getRow(TestCaseRowNo);
			XSSFCell cell2 = Row1.getCell(0);
//			System.out.println("col no"+TotalColNO+cell2);
			break;
		}
	}
		for (int j = (TestCaseRowNo+1); j < TotalNoRows; j++) {
			XSSFRow Row =WSName.getRow(j);
			XSSFCell cell = Row.getCell(ColCheckMode);
			//System.out.println("header no "+(HeaderRowNo+1));
//			System.out.println("run mode"+cell.getStringCellValue());
			if (cell.getStringCellValue().trim().equalsIgnoreCase(TestCaseName)){
				count=j;
			}

	}
//	System.out.println("count No "+count);
	
	XSSFRow HeaderRow = WSName.getRow(TestCaseRowNo);
	for (int i = 0; i < TotalColNO; i++) {
	XSSFCell cell=	HeaderRow.getCell(i);
	
	if (cell.getStringCellValue().trim().equalsIgnoreCase(colName)) {
//		System.out.println("cell Value"+cell.getStringCellValue());
		countNo=i;
	}
	}
	
	if (countNo==-1) {
//		System.out.println("Count No Nullllllllllllll");
		return null;
	}
	
	XSSFRow Row = WSName.getRow(count);
	XSSFCell cell = Row.getCell(countNo);
	String value = cell.getStringCellValue();
		return value;
	
	}


/**
 * This method is used for write execution status in test data  
 * @param SheetName Sheet name -- Test Data
 * @param HeaderOrTastCaseName Header and Test case name are same in  Test Data sheet   
 * @param colName Pass/Fail/Skip
 * @param rowNumber Row number start with 0
 * @param Result write result in cell as per SheatName->TestCaseName->Status update column name
 * @return 
 */

public boolean writeExciquteResultInTestData(String SheetName, String HeaderOrTastCaseName, String colName, int rowNumber, String Result){

	int ColHeaderNo=0;
	int HeaderRowNo = 0;
	int count=0;
	int countNo=-1;
	int TotalColNO = 0;
	
	
	WSName=WB.getSheet(SheetName);
	int TotalNoRows = WSName.getLastRowNum()+1;
//	System.out.println("Total Rows"+TotalNoRows);
	if (TotalNoRows==0) {
		//return null;
//		System.out.println("null");
	}
	
	for (int i = 0; i < TotalNoRows; i++) {
		XSSFRow Row = WSName.getRow(i);
		if (Row==null) {
			continue;
		}
		XSSFCell cell = Row.getCell(ColHeaderNo);
//		System.out.println("Header Name "+cell.getStringCellValue());
		if (cell.getStringCellValue().trim().equalsIgnoreCase(HeaderOrTastCaseName)) {
			HeaderRowNo = i+1;
			TotalColNO = WSName.getRow(HeaderRowNo).getLastCellNum();
//			System.out.println("header row no "+HeaderRowNo);
			XSSFRow Row1 = WSName.getRow(HeaderRowNo);
			XSSFCell cell2 = Row1.getCell(0);
//			System.out.println("col no"+TotalColNO+cell2);
			break;
		}
	}
	
//	System.out.println("count No "+count);
	
	XSSFRow HeaderRow = WSName.getRow(HeaderRowNo);
	for (int i = 0; i < TotalColNO; i++) {
	XSSFCell cell=	HeaderRow.getCell(i);
	
	if (cell.getStringCellValue().trim().equalsIgnoreCase(colName)) {
//		System.out.println("cell Value"+cell.getStringCellValue());
		countNo=i;
	}
	}
	
	if (countNo==-1) {
		return false;
	}
	
XSSFRow Row=WSName.getRow(HeaderRowNo+rowNumber);
XSSFCell cell=Row.getCell(countNo);
if (cell==null)
	cell = Row.createCell(HeaderRowNo+rowNumber);
cell.setCellValue(Result);

try {
	FOS=new FileOutputStream(FileLoation);
	WB.write(FOS);
	FOS.close();
} catch (FileNotFoundException e) {
		e.printStackTrace();
		return false;
} catch (IOException e) {
		e.printStackTrace();
		return false;
}
return true;
	
	
}


/**
 * This method is used for write execution status in test case list  
 * @param SheetName Sheet name -- Test Case list
 * @param Header  Header name  -- Test Case list
 * @param TastCaseName Test Case Name of execution 
 * @param colName Pass/Fail/Skip
 * @param Result write result in cell as per SheatName->Header -> TestCaseName->Status update column name
  * @return
 */
public boolean writeExciquteResultInTestCase(String SheetName, String Header, String TastCaseName, String colName,  String Result){

	int ColHeaderNo=0;
	int HeaderRowNo = 0;
	int count=0;
	int countNo=-1;
	int TotalColNO = 0;
	
	WSName=WB.getSheet(SheetName);
	int TotalNoRows = WSName.getLastRowNum()+1;
//	System.out.println("Total Rows"+TotalNoRows);
	if (TotalNoRows==0) {
		//return null;
//		System.out.println("null");
	}
	
	for (int i = 0; i < TotalNoRows; i++) {
		XSSFRow Row = WSName.getRow(i);
		if (Row==null) {
			continue;
		}
		XSSFCell cell = Row.getCell(ColHeaderNo);
//		System.out.println("Header Name "+cell.getStringCellValue());
		if (cell.getStringCellValue().trim().equalsIgnoreCase(Header)) {
			HeaderRowNo = i+1;
			TotalColNO = WSName.getRow(HeaderRowNo).getLastCellNum();
//			System.out.println("header row no "+HeaderRowNo);
			XSSFRow Row1 = WSName.getRow(HeaderRowNo);
			XSSFCell cell2 = Row1.getCell(0);
//			System.out.println("col no"+TotalColNO+cell2);
			break;
		}
	}

	for (int j = (HeaderRowNo+1); j < TotalNoRows; j++) {
			XSSFRow Row =WSName.getRow(j);
			XSSFCell cell = Row.getCell(ColHeaderNo);
			//System.out.println("header no "+(HeaderRowNo+1));
//			System.out.println("run mode"+cell.getStringCellValue());
			if (cell.getStringCellValue().trim().equalsIgnoreCase(TastCaseName)) {
				count=j;
			}

	}
//	System.out.println("count No "+count);
	
	XSSFRow HeaderRow = WSName.getRow(HeaderRowNo);
	for (int i = 0; i < TotalColNO; i++) {
	XSSFCell cell=	HeaderRow.getCell(i);
	
	if (cell.getStringCellValue().trim().equalsIgnoreCase(colName)) {
//		System.out.println("cell Value"+cell.getStringCellValue());
		countNo=i;
	}
	}
	
	if (countNo==-1) {
//		System.out.println("Count No Nullllllllllllll");
		//return null;
	}
	
XSSFRow Row=WSName.getRow(count);
XSSFCell cell=Row.getCell(countNo);
if (cell==null)
	cell = Row.createCell(countNo);
cell.setCellValue(Result);

try {
	FOS=new FileOutputStream(FileLoation);
	WB.write(FOS);
	FOS.close();
} catch (FileNotFoundException e) {
		e.printStackTrace();
		return false;
} catch (IOException e) {
		e.printStackTrace();
		return false;
}
return true;
	
	
}


}

