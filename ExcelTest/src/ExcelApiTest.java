import java.io.FileInputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelApiTest {
	public FileInputStream fis = null;
    public XSSFWorkbook wb = null;
    public XSSFSheet sh = null;
    public XSSFRow row = null;
    public XSSFCell cell = null;
    
    public ExcelApiTest(String xlFilePath) throws Exception
    {
        fis = new FileInputStream(xlFilePath);
        wb = new XSSFWorkbook(fis);
        fis.close();
    }
    
    public String getCellData(String SheetName, String colName, int rowNum) {
    	try {
    	int colnum=-1;
    	sh=wb.getSheet(SheetName);
    	row = sh.getRow(0);
    	for(int i=0;i<row.getLastCellNum();i++) {
    		if(row.getCell(i).getStringCellValue().trim().equals(colName.trim())) {
    			colnum=i;
    		}
    	}
    	row = sh.getRow(rowNum-1);
    	cell = row.getCell(colnum);
    	
    	if(cell.getCellTypeEnum()==CellType.STRING) {
    		return cell.getStringCellValue();
    	}else if(cell.getCellTypeEnum()==CellType.NUMERIC || cell.getCellTypeEnum()==CellType.FORMULA) {
    		String cellValue = String.valueOf(cell.getNumericCellValue());
    		if(HSSFDateUtil.isCellDateFormatted(cell)) {
    			DateFormat df = new SimpleDateFormat("DD/MM/YY");
    			Date date = cell.getDateCellValue();
    			cellValue = df.format(date);
    		}
    		return cellValue;
    	}else if(cell.getCellTypeEnum()==CellType.BLANK) {
    		return "";
    	}else {
    		return String.valueOf(cell.getBooleanCellValue());
    	}
    	}catch(Exception e) {
    		e.printStackTrace();
    		return "row "+rowNum+" or column "+colName +" does not exist  in Excel";
    	}
    }
	public static void main(String[] args) throws Exception {
		ExcelApiTest eat = new ExcelApiTest("C:/Users/NavPooja/Desktop/ExcelTest.xlsx");        
        System.out.println(eat.getCellData("Sheet1","Username",2));
        System.out.println(eat.getCellData("Sheet1","Password",2));
        System.out.println(eat.getCellData("Sheet1","DateCreated",2));
        System.out.println(eat.getCellData("Sheet1","NoOfAttempts",2));	
	}
}
