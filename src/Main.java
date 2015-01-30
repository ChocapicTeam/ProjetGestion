import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import jxl.read.biff.BiffException;

 
public class Main {

	public static void main(String[] args) throws BiffException, IOException, InvalidFormatException   {
		
		String name = "resultat.xls";
		
		try {
			TestExcel.readXLSWithPOI(name);
		}
		catch (org.apache.poi.hssf.OldExcelFormatException e) {
			TestJXL.readXLSWithJXL(name);
		}		
	}

}
