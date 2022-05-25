package XmlWriteReadFile;


//assignment: read all the data from all the existing sheets in the given xl file
import java.io.File;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.Map;
import java.util.Properties;
import java.util.Set;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//
public class ReadWriteXlData1 {
	
public static void readDataFromXlsxCell(int rowData,int colData,String sheetName,String path) throws IOException {
	FileInputStream fs=new FileInputStream(new File(path));
	XSSFWorkbook workbook=new XSSFWorkbook(fs);// .xlsx
	XSSFSheet sheet=workbook.getSheet(sheetName);
	XSSFRow row=sheet.getRow(rowData);
	XSSFCell cell=row.getCell(colData);
	if(cell.getCellType()== CellType.NUMERIC)
		System.out.println(cell.getNumericCellValue());
	else if(cell.getCellType()== CellType.STRING)
		System.out.println(cell.getStringCellValue());
}

public static void writeDataToXlsxCell(int rowData,int colData,String sheetName,String data, String path) throws IOException {
	FileInputStream fs=new FileInputStream(new File(path));
	XSSFWorkbook workbook=new XSSFWorkbook(fs);// .xlsx
	XSSFSheet sheet=workbook.getSheet(sheetName);		
	XSSFRow row=sheet.getRow(rowData);
	XSSFCell cell=row.getCell(colData);
	cell.setCellValue(data);
	FileOutputStream fo=new FileOutputStream(new File(path));
	workbook.write(fo);
	fo.flush();
	fo.close();
}
public static void writeSingleDataToNewXlsxFile(String sheetName, String path,String data) throws IOException {
	
	
	XSSFWorkbook workbook = new XSSFWorkbook();// .xlsx
	XSSFSheet sheet = workbook.createSheet(sheetName);
	XSSFRow row=sheet.createRow(0);
	XSSFCell cell=row.createCell(0);
	cell.setCellValue(data);
	FileOutputStream fo=new FileOutputStream(new File(path));
	workbook.write(fo);
	fo.flush();
	fo.close();
	
}


public static void ReadDataFromXlsxSheet(String sheetName,String path) throws IOException {
	FileInputStream fs=new FileInputStream(new File(path));
	XSSFWorkbook workbook=new XSSFWorkbook(fs);// .xlsx

	XSSFSheet sheet=workbook.getSheet(sheetName);
	Iterator<Row> rows=sheet.iterator();
	while(rows.hasNext()) {
		Row oneRow=rows.next();
		Iterator<Cell> cells=oneRow.cellIterator();
		while(cells.hasNext()) {
			Cell oneCell= cells.next();
			if(oneCell.getCellType()== CellType.NUMERIC)
				System.out.print(oneCell.getNumericCellValue()+"    ");
			else if(oneCell.getCellType()== CellType.STRING)
				System.out.print(oneCell.getStringCellValue()+"   ");
		}
		System.out.println();
		
	}
	
	
}

public static void ReadFromPropertyFile(String pathproperty) throws IOException
{
	FileInputStream fs=new FileInputStream(new File(".//src//Resource//Data1.properties"));
	Properties p=new Properties();
	p.load(fs);
	System.out.println(p.getProperty("batch"));
	
	FileOutputStream fo=new FileOutputStream(new File(".//src//Resource//Data1.properties"));
	Properties p1=new Properties();
	p1.setProperty("batch","Dec22");
	p1.store(fo, null);
	fo.close();
	
}
public static void main(String[] args) throws IOException {
	String path=".//src//Resource//Dec22Exel.xlsx";
	String path1=".//src//Resource//test.xlsx";
	String newFile=".//src//Resource//newXlSheet.xlsx";
	//readDataFromXlsxCell(1,0,"Sheet1",path);
	//readDataFromXlsxCell(1,0,"Sheet1",path1);
	//readDataFromXlsxCell(1,2,"Sheet1",path);
	
	//ReadDataFromXlsxSheet("Sheet1",path);
	
	writeDataToXlsxCell(2,2,"Sheet11","Selenium11",path);
	
	
	// TODO Auto-generated method stub
			
	writeSingleDataToNewXlsxFile("BatchStatus",newFile,"Dec21");
	System.out.println("completed");
	
		
}

}