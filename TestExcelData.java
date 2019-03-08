package ExcelDataReading;

import java.io.FileInputStream;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class TestExcelData
{
	public void Excel1(String FilePath)
	 {		
		try{ 
		 FileInputStream fs = new FileInputStream(FilePath);
		 
		 XSSFWorkbook workbook= new XSSFWorkbook(fs);
		 XSSFSheet sheet= workbook.getSheetAt(0);
		 int lastrow = sheet.getLastRowNum();
		 XSSFRow row = sheet.getRow(0);

		String val = row.getCell(0).getStringCellValue();
		System.out.println(val);
		 
		 fs.close();
		 
		} catch (Exception objexc)
		{
			objexc.printStackTrace();
		}
	}
	
public static void main(String[] args) 
{
	TestExcelData obj = new TestExcelData();
	obj.Excel1("D:\\Sandip Selenium setups\\TestData\\TestData.xlsx");
	
}
}