import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class excelRead {

	public static void main(String[] args) throws IOException {
		FileInputStream fis = new FileInputStream("C:/Users/NavPooja/Desktop/PythonTest.xlsx");
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sh = wb.getSheetAt(0);
		XSSFRow row = sh.getRow(0);
		
		int col_num = -1;
		
		for(int i=0;i<row.getLastCellNum();i++) {
			if(row.getCell(i).getStringCellValue().trim().equals("EMP_ID")) {
				col_num=i;
			}
		}
		row = sh.getRow(1);
		XSSFCell cell = row.getCell(col_num);
		String value = cell.getStringCellValue();
		System.out.println("Value of the excel cell is: "+value);
	}
}
