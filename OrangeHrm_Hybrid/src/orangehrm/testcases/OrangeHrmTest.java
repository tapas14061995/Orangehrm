package orangehrm.testcases;

import java.io.IOException;

import org.testng.annotations.Test;

import orangehrm.library.Employee;
import orangehrm.library.LoginPage;
import utils.AppUtils;
import utils.XLUtils;

public class OrangeHrmTest extends AppUtils
{
	String tcfile= "C:\\Users\\91828\\OneDrive\\Desktop\\OrangeHrm.xlsx";
	String tcsheet="TestCases";
	String tssheet="TestSteps";
	@Test
	public void checkOrangeHrm() throws IOException
	{
		
		String tcexeflag,tcid,tstcid,keyword;
		String uid,pwd;
		String fname,lname;
		String stepres,tcres;
		boolean res = false;
		
		int tccount = XLUtils.getRowCount(tcfile, tcsheet);
		int tscount = XLUtils.getRowCount(tcfile, tssheet);
		LoginPage lp =new LoginPage();
		Employee emp =new Employee();
		
		for(int i=1;i<=tccount;i++)
		{
			tcexeflag = XLUtils.getStringData(tcfile, tcsheet, i, 2);
			if(tcexeflag.equalsIgnoreCase("Y"))
			{
				tcid = XLUtils.getStringData(tcfile, tcsheet, i, 0);
				for(int j=1;j<=tscount;j++)
				{
					tstcid=XLUtils.getStringData(tcfile, tssheet, j, 0);
					if(tstcid.equalsIgnoreCase(tcid))
					{
						keyword=XLUtils.getStringData(tcfile, tssheet, j, 4);
						switch (keyword.toLowerCase())
						{
						case "adminlogin":
							uid =XLUtils.getStringData(tcfile, tssheet, j, 5);
							pwd =XLUtils.getStringData(tcfile, tssheet, j, 6);
							lp.login(uid, pwd);
							res=lp.isAdminModuleDisplayed();
							break;
							
						case "logout":
							res=lp.logout();
						    break; 
						   	
						case "invalidlogin":
							uid = XLUtils.getStringData(tcfile, tssheet, j, 5);
							pwd = XLUtils.getStringData(tcfile, tssheet, j, 6);
							lp.login(uid, pwd);
							res = lp.isErrMsgDisplayed();
							break;
							
						case "empreg":
							fname=XLUtils.getStringData(tcfile, tssheet, j, 5);
							lname=XLUtils.getStringData(tcfile, tssheet, j, 6);
							res=emp.addEmployee(fname, lname);
							break;
							
						}
						
						//code to update step result
							
						if(res)
						{
							stepres="pass";
							XLUtils.setData(tcfile, tssheet, j, 3, stepres);
							XLUtils.fillGreenColor(tcfile, tssheet, j, 3);
						}else
						{
							stepres="fail";
							XLUtils.setData(tcfile, tssheet, j, 3, stepres);
							XLUtils.fillRedColor(tcfile, tssheet, j, 3);
						}
						
						//code to update TestCase result
						
						tcres = XLUtils.getStringData(tcfile, tcsheet, i, 3);
						if(!tcres.equalsIgnoreCase("fail"))
						{
							XLUtils.setData(tcfile, tcsheet, i, 3, stepres);
						}
						
						//code to fill TcRes color
						
						tcres = XLUtils.getStringData(tcfile, tcsheet, i, 3);
						if(tcres.equalsIgnoreCase("pass"))
						{
							XLUtils.fillGreenColor(tcfile, tcsheet, i, 3);
						}else
						{
							XLUtils.fillRedColor(tcfile, tcsheet, i, 3);
						}
					}
				}				
			}else
			{
				XLUtils.setData(tcfile, tcsheet, i, 3, "Blocked");
				XLUtils.fillRedColor(tcfile, tcsheet, i, 3);
			}			
		}
			
	}
}
