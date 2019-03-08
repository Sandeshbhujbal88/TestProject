package ExcelDataReading;

import java.io.FileInputStream;
import java.util.ArrayList;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class testexcel3 
{
	public ArrayList Excel1()
	 {
		ArrayList<E> <String> onjArList= new ArrayList <String>();
		 
		try{
			 
		 FileInputStream fs = new FileInputStream("G:\Selenium testing\\TestData.xlsx");
		 XSSFWorkbook workbook= new XSSFWorkbook(fs);
		 XSSFSheet sheet= workbook.getSheetAt(0);
		
		 int lastrow =sheet.getLastRowNum();
		 for (int a=0; a<=lastrow; a++)
		 {
			 XSSFRow row = sheet.getRow(a);
			 int c= row.getLastCellNum();
			 for (int d=0; d<c; d++)
			 {
				 String val= row.getCell(d).getStringCellValue();
				 //System.out.print(val+" ");
				 
				 onjArList.add(val);
			 }
			 System.out.println();
		 }
		 fs.close();
		 }
		 catch(Exception objExc)
		 {
			 objExc.printStackTrace();
		 }
		return onjArList;
	}
	 
	 public static void main(String[]args)
	 {
		 testexcel3 obj = new testexcel3();
		 obj.Excel1();
		 
	 }
}
