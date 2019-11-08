import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class APB {
	private static final long NumericCellValue = 0;

	public static void main(String[] args) throws Throwable {
		File f=new File("C:\\Users\\admin\\Desktop\\Greens\\LAP\\TESTFILE\\MERIN.xlsx");
		FileInputStream s=new FileInputStream(f);
		Workbook w= new XSSFWorkbook(s);
		Sheet s1 = w.getSheet("Sheet1");
		for (int i = 0; i <s1.getPhysicalNumberOfRows(); i++) {
			Row r = s1.getRow(i);
			for (int j = 0; j <r.getPhysicalNumberOfCells(); j++) {
				Cell c = r.getCell(j);
				int type = c.getCellType();
				if (type==1) {
					String scv = c.getStringCellValue();
					System.out.println(scv);
					
				}
				else if (type==0) {
					if (DateUtil.isCellDateFormatted(c)) {
						Date dcv = c.getDateCellValue();
						SimpleDateFormat sdf=new SimpleDateFormat("dd-mmm-yy");
						String s2 = sdf.format(dcv);
						System.out.println(s2);
						
					}
					else{
						double ncv = c.getNumericCellValue();
						//convert double into long
						long l=(long)NumericCellValue;
						//convert long into string
						String l2 = String.valueOf(l);
						System.out.println(l2);
					}
					
				}
			}
			
		}

		
			
		
		
	}

}
