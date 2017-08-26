package junit2;

import java.io.File;

import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import jxl.Sheet;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

public class readandwrite 
{



	public static void main(String[] args)throws Exception
	
	{
		System.setProperty("webdriver.chrome.driver", "D:\\tet\\Selenium_Demp\\chromedriver_win32\\chromedriver.exe");
		
		ChromeDriver Driver=new ChromeDriver();
		WebDriverWait Wait= new WebDriverWait(Driver, 40);
	
		Driver.get("http://classroom:90/qahrm");
		Driver.manage().window().maximize();
		
		File fis=new File("D:\\143.xls");
		Workbook Objwb=Workbook.getWorkbook(fis);
		Sheet InputSheet=Objwb.getSheet(0);
		WritableWorkbook wwb=Workbook.createWorkbook(new File("D:\\res.xls"));
		WritableSheet OutputSheet=wwb.createSheet("HRM", 0);
		Label l1=new Label(0,0,"username");
		Label l2=new Label(1,0,"password");
		Label k1=new Label(2,0,"Result");
	    OutputSheet.addCell(l1);
	    OutputSheet.addCell(l2); 
	    OutputSheet.addCell(k1);
	    Label r1;
	    Label r2;
	    Label r3;
	    int Rcount=InputSheet.getRows();
	    System.out.println("Rcount");
	    for(int i=1;i<Rcount;i++)
	    {
	    	System.out.println("Iteration no:"+i);
	    	
	    	
	    	WebElement ObjUN=Wait.until(ExpectedConditions.presenceOfElementLocated(By.name("txtUserName")));
                ObjUN.clear();	    
	        ObjUN.sendKeys(InputSheet.getCell(0,i).getContents());
	        WebElement ObjPWD=Wait.until(ExpectedConditions.presenceOfElementLocated(By.name("txtPassword")));
	            ObjPWD.clear();
	        ObjPWD.sendKeys(InputSheet.getCell(1,i).getContents() );
	        WebElement Objclick=Wait.until(ExpectedConditions.presenceOfElementLocated(By.name("Submit")));
	         Objclick.click();
	         Thread.sleep(2000);
	         String res="passed";
	         String res1="failed";
	         if(Driver.getTitle().equals("OrangeHRM") )
	         {
	        	 System.out.println("login sucessfull");
	         
	         r1=new Label(0,i, InputSheet.getCell(0, i).getContents());
	         r2=new Label(1,i, InputSheet.getCell(1, i).getContents());
	         r3=new Label(2,i,res);
	         OutputSheet.addCell(r1);
	         OutputSheet.addCell(r2);
	         OutputSheet.addCell(r3);
	         
	      Driver.findElement(By.linkText("Logout")).click();
	      Thread.sleep(3000);
	    	System.out.println("Logout is done");
	    	}
	    else
	    {
	    	System.out.println("logout is failed");
	    	 r1=new Label(0,i, InputSheet.getCell(0, i).getContents());
	         r2=new Label(1,i, InputSheet.getCell(1, i).getContents());
	         r3=new Label(2,i,res);
	         OutputSheet.addCell(r1);
	         OutputSheet.addCell(r2);
	         OutputSheet.addCell(r3);
	         
	      Driver.findElement(By.name("clear")).click();
	    	wwb.write();
	    	wwb.close();
	    }
	    }
	    Driver.close();
	    Driver.quit();
	}
	    	
	         
	        
	    
	    
	    
	    
	    
	    
	    
	    
	    
	    
	    
	    
	    
	    
	    
	
}
