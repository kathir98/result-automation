package pondiuniresult;


import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.*;
import java.util.Scanner;

import org.apache.commons.io.FileUtils;


public class pondiuniscreenshot {
	public static void main(String[] args) throws InterruptedException, AWTException, IOException {
		Scanner s=new Scanner(System.in);
		
//		// TO get url
//    	System.out.println("Enter the URL");
//    	String baseurl = s.next();
//    	System.out.println();
		
		String baseurl="http://exam.pondiuni.edu.in/oresults/";
		
		// To get the semester number 
    	System.out.println("Enter The Semester To Which You Want To Get Screenshot Of Results");
    	System.out.println("(i.e) First | Second | Third | Fourth | Fifth | Sixth | Seventh | Eight");
    	String sem = s.next();
    	String firstLetter = sem.substring(0,1).toUpperCase();
        String restLetters = sem.substring(1).toLowerCase();
        String semesterno =firstLetter + restLetters;
        System.out.println();
        
        // TO get the source of the register number
    	System.out.println("Enter The FILE LOCATION and FILENAME OF the input file containing REGISTER NUMBERS");
    	System.out.println("COPY AND PASTE TO AVOID ERRORS");
    	String sourcelocation = s.next();
    	System.out.println();
        
//        // TO get the destination folder to save screenshots
//    	System.out.println("Enter The FILE LOCATION to save the DESTINATION FILES");
//    	System.out.println("COPY AND PASTE TO AVOID ERRORS");
//    	String destinationlocation = s.next();
//    	System.out.println();
    	
		// Setting webdriver and other stuffs
		System.setProperty("webdriver.chrome.driver","F:\\selenium\\chromedriver.exe");
		WebDriver driver = new ChromeDriver();
		JavascriptExecutor js = (JavascriptExecutor) driver;
		Robot robot = new Robot();
		driver.manage().window().maximize();
		
		// Passing url to the driver
		driver.get(baseurl);
		
		BufferedReader reader = null;
		try {

			// Reading roll number as input from file
			reader = new BufferedReader(new FileReader(sourcelocation));
			String line = reader.readLine();
			while(line != null)
			{
				try
				{
				// Finding element and passing values
				WebElement registerno = driver.findElement(By.id("reg_no"));
				registerno.sendKeys(line);
				WebElement semester = driver.findElement(By.id("exam"));
				semester.sendKeys(semesterno);
				WebElement submit = driver.findElement(By.xpath("//*[@id=\"print_app_form\"]"));
				submit.click(); 
				
				// It is used to delay the zoomout operation. If it is not used the webpage is zoomedout earlier					
				WebDriverWait wait = new WebDriverWait(driver, 60); 
				wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("/html/body/div[1]/div/div[3]/div/div[2]/div[1]/div[2]/table[1]/tbody/tr[2]/td"))); 
				
				// To reduce zoom i.e zoomout
				robot.keyPress(KeyEvent.VK_CONTROL);
				robot.keyPress(KeyEvent.VK_SUBTRACT);
				robot.keyRelease(KeyEvent.VK_SUBTRACT);
				robot.keyRelease(KeyEvent.VK_CONTROL);
			
				// To scroll down to the appropriate position
				WebElement crop = driver.findElement(By.className("result_table_header"));
				js.executeScript("arguments[0].scrollIntoView();", crop);
			
				// To take screenshot
				File screenshot = ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
				FileUtils.copyFile(screenshot, new File("screenshots\\"+line+".png"));
				
				// To increase zoom so that to avoid zoomin again in next iteration
				robot.keyPress(KeyEvent.VK_CONTROL);
				robot.keyPress(KeyEvent.VK_ADD);
				robot.keyRelease(KeyEvent.VK_ADD);
				robot.keyRelease(KeyEvent.VK_CONTROL);
				
				// To navigate back and try next input
				driver.navigate().back();
				line = reader.readLine();
				registerno.clear();
				semester.clear();
			    }
				catch(Exception e)
				{
					System.out.println("exception occured :"+e);
					String alert="org.openqa.selenium.UnhandledAlertException";
					String excep=e.toString();  
					if(excep.startsWith(alert))
					{
					line = reader.readLine();
					}
					WebElement registerno = driver.findElement(By.id("reg_no"));
					registerno.clear();
				}
		}
		}
		catch(Exception e)
		{
			System.out.println(e);
		}

		reader.close();
		driver.quit();
		System.out.println("The Screenshots Taken");
		

	}

}
