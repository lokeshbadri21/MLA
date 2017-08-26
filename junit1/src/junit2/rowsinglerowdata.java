package junit2;

import java.io.File;
import java.io.IOException;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

public class rowsinglerowdata 
{
	 public static void main(String[] args) throws BiffException, IOException 
	    {
	        File f1=new File("d:\\123.xls");
	        Workbook WB=Workbook.getWorkbook(f1);
	        Sheet s1=WB.getSheet("QAP");
	        int i=1;
	        String gBrow=s1.getCell(0, i).getContents();
	        String gURL=s1.getCell(1, i).getContents();
	        String gUN=s1.getCell(2, i).getContents();
	        String gPWD=s1.getCell(3, i).getContents();
	        String gEMAIL=s1.getCell(4, i).getContents();
	        System.out.println(gBrow);
	        System.out.println(gURL);
	        System.out.println(gUN);
	        System.out.println(gPWD);
	        System.out.println(gEMAIL);
	  
	  }

}
