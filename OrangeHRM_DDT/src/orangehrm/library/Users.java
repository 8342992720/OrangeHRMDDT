package orangehrm.library;

import java.util.Iterator;
import java.util.List;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.Select;
import org.yaml.snakeyaml.events.Event.ID;

import utils.AppUtils;

public class Users extends AppUtils
{
	public boolean userReg(String role,String ename,String uname,String pwd) throws InterruptedException
	{
		driver.findElement(By.linkText("Admin")).click();
		driver.findElement(By.linkText("User Management")).click();
		driver.findElement(By.linkText("Users")).click();
		driver.findElement(By.id("btnAdd")).click();
		
		Select rolelist=new Select(driver.findElement(By.id("systemUser_userType")));
		rolelist.selectByVisibleText(role);
		
		driver.findElement(By.id("systemUser_employeeName_empName")).sendKeys(ename);
		driver.findElement(By.id("systemUser_userName")).sendKeys(uname);
		driver.findElement(By.id("systemUser_password")).sendKeys(pwd);
		driver.findElement(By.id("systemUser_confirmPassword")).sendKeys(pwd);
		
		Thread.sleep(5000);
		
		driver.findElement(By.id("btnSave")).click();
	
	
		WebElement table=driver.findElement(By.id("resultTable"));
		List<WebElement> rows,cols;
		rows=table.findElements(By.tagName("tr"));
		
		boolean userpresent= false;
		for(int i=1;i<rows.size();i++)
		{
			cols=rows.get(i).findElements(By.tagName("td"));
			if(cols.get(1).getText().equals(uname))
			{
				userpresent= true;
				break;
			}
			
		}return userpresent;
		
		}
			
	}
	
	
	
	
	
	
	


