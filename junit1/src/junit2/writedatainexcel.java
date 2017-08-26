package junit2;

import java.io.File;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

public class writedatainexcel 
{
	public static void main(String[] args)throws Exception
	{
		WritableWorkbook wwb=Workbook.createWorkbook(new File("d:\\loke.xls"));
		WritableSheet wb=wwb.createSheet("HRM",0);
		WritableSheet ws1=wwb.createSheet("RES",1);
		Label k1=new Label(0,0,"username");
		Label k2=new Label(1,0,"password");
		Label r1=new Label(0,1,"qaplanet1");
      	Label r2=new Label(1,1,"admin");
      	wb.addCell(k1);
      	wb.addCell(k2);
      	wb.addCell(r1);
      	wb.addCell(r2);
      	wwb.write();
      	wwb.close();

	
	}
}
