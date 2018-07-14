package test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Exceldriven {

	public static FileInputStream fis;
	public static XSSFWorkbook wb;
	public static XSSFSheet ws;
	public static XSSFRow rw;
	public static XSSFCell cl;
	
	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		 
		
		String value = getCelldata(1,2);
		System.out.println(value);
		String value1 = setCelldata("Auto",3,2);
		System.out.println(value1);
	}
public static String getCelldata(int row, int col) throws IOException
{
	fis = new FileInputStream("C:\\jas-hadoop\\selenium\\data_driven_1.xlsx");
	wb = new XSSFWorkbook(fis);
	 ws = wb.getSheet("script");
	 rw = ws.getRow(3);
	 cl = rw.getCell(2);
	return cl.getStringCellValue();
}

public static String setCelldata(String text, int row, int col) throws IOException
{
	fis = new FileInputStream("C:\\jas-hadoop\\selenium\\data_driven_1.xlsx");
	wb = new XSSFWorkbook(fis);
	 ws = wb.getSheet("script");
	 rw = ws.getRow(3);
	 cl = rw.getCell(2);
	 cl.setCellValue(text);// cl.setCellValue("hi") we can hard code in this reusable code, but its better this way
	
	FileOutputStream fos = new FileOutputStream("C:\\\\jas-hadoop\\\\selenium\\\\data_driven_1.xlsx"); 
	wb.write(fos); 
	fos.close();
	return cl.getStringCellValue();
}

}
