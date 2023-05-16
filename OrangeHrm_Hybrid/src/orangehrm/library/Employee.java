package orangehrm.library;

import java.util.List;

import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;

import utils.AppUtils;

public class Employee extends AppUtils
{
	public boolean addEmployee(String fname, String lname)
	{
		driver.findElement(By.linkText("PIM")).click();
		driver.findElement(By.linkText("Add Employee")).click();
		
		driver.findElement(By.name("firstName")).sendKeys(fname);
		driver.findElement(By.id("lastName")).sendKeys(lname);
		
		String empid=driver.findElement(By.id("employeeId")).getAttribute("value");
		
		driver.findElement(By.id("btnSave")).click();
		
		driver.findElement(By.linkText("Employee List")).click();
		driver.findElement(By.id("empsearch_id")).sendKeys(empid);
		driver.findElement(By.id("searchBtn")).click();
		
		WebElement restbl = driver.findElement(By.id("resultTable"));
		List<WebElement> rows= restbl.findElements(By.tagName("tr"));
		
		boolean isempexist = false;
		for(int i=1;i<rows.size();i++)
		{
			List<WebElement> cols =  rows.get(i).findElements(By.tagName("td"));
			if(cols.get(i).getText().equals(empid))
			{
				isempexist=true;
				break;
			}
		}
		return isempexist;
	}
	
	public boolean addEmployee(String fname, String mname, String lname)
	{
		driver.findElement(By.linkText("PIM")).click();
		driver.findElement(By.linkText("Add Employee")).click();
		
		driver.findElement(By.name("firstName")).sendKeys(fname);
		driver.findElement(By.name("middleName")).sendKeys(mname);
		driver.findElement(By.id("lastName")).sendKeys(lname);
		
		String empid=driver.findElement(By.id("employeeId")).getAttribute("value");
		
		driver.findElement(By.id("btnSave")).click();
		
		driver.findElement(By.linkText("Employee List")).click();
		driver.findElement(By.id("empsearch_id")).sendKeys(empid);
		driver.findElement(By.id("searchBtn")).click();
		
		WebElement restbl = driver.findElement(By.id("resultTable"));
		List<WebElement> rows= restbl.findElements(By.tagName("tr"));
		
		boolean isempexist = false;
		for(int i=1;i<rows.size();i++)
		{
			List<WebElement> cols =  rows.get(i).findElements(By.tagName("td"));
			if(cols.get(i).getText().equals(empid))
			{
				isempexist=true;
				break;
			}
		}
		return isempexist;
	}	
}
