package generic;

import java.io.FileInputStream;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Excel {

	public static String getData(String path,String sheet,int r,int c) {
		String v="";
			try {
				System.out.println(path);
				System.out.println(sheet);
				System.out.println(r);
				System.out.println(c);
				Workbook wb = WorkbookFactory.create(new FileInputStream(path));
				v = wb.getSheet(sheet).getRow(r).getCell(c).toString();
				wb.close();
				System.out.println("Exceldata" +v); 
			}
			catch (Exception e) {
				e.printStackTrace();
			}
		return v;
	}
	
	public static int getRowCount(String path,String sheet) {
		int v=0;
			try {
				Workbook wb = WorkbookFactory.create(new FileInputStream(path));
				v = wb.getSheet(sheet).getLastRowNum();
				wb.close();
			}
			catch (Exception e) {
				
			}
		return v;
	}

}
