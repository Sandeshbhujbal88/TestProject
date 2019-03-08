package ExcelDataReading;

import java.io.FileInputStream;
import java.util.ArrayList;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class rvexcelTest
{
public ArrayList rvexcel()
{
	ArrayList<String> testlist = new ArrayList<String>();
try {
	FileInputStream fs= new FileInputStream("G:\\TestData.xlsx");
    XSSFWorkbook workbook= new XSSFWorkbook(fs);
    XSSFSheet sheet= workbook.getSheetAt(0);
    
    int lastrow=sheet.getLastRowNum();
    
    for ( int a=0; a<lastrow; a++)
    {
    		XSSFRow row= sheet.getRow(a);
    		int lastcell= row.getLastCellNum();
    	
    			String temp = row.getCell(0).getStringCellValue();
    		//	System.out.print(temp +" ");
    			testlist.add(temp)
    		}
    		System.out.println();
    }
  fs.close();
}
catch(Exception objExc)
{
	 objExc.printStackTrace();
}
return testlist();
}

public static void main(String[]args)
{
	rvexcelTest obj = new rvexcelTest();
	obj.rvexcel();
}
}
