package XmlWriteReadFile;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadWriteXlData {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		//connect file to the sheet
		FileInputStream fs=new FileInputStream(new File(".//src//Resource//Dec22Exel.xlsx"));
		//get the work book from the file
		XSSFWorkbook workbook=new XSSFWorkbook(fs);// .xlsx
		//get the worksheet from the workBook
		XSSFSheet sheet=workbook.getSheet("Sheet1");
		int row1=sheet.getLastRowNum();
		int colm1=sheet.getRow(1).getLastCellNum();
		XSSFRow row=sheet.getRow(0);
		XSSFCell cell=row.getCell(0);
		if(cell.getCellType()== CellType.NUMERIC)
			System.out.println(cell.getNumericCellValue());
		else if(cell.getCellType()== CellType.STRING)
			System.out.println(cell.getStringCellValue());
		
		
		
		
			
	}

}