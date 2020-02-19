package pondiuniresult;
// adding header files
import java.awt.AWTException;
import java.awt.Robot;
import java.io.BufferedReader;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Scanner;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class pondiuniexcel {
	public static void main(String ar[]) throws AWTException, IOException
	{
		Scanner s=new Scanner(System.in);
		
//		// TO get url
//		System.out.println("Enter the URL");
//		String baseurl = s.next();
//		System.out.println();
		
		String baseurl="http://exam.pondiuni.edu.in/oresults/";
		XSSFWorkbook workbook = new XSSFWorkbook();
		FileOutputStream fileOut=new FileOutputStream("results.xlsx");	//Path of Output File
		int index=1,lineNo=0;
		XSSFSheet sheet = workbook.createSheet("Result");
		ArrayList<String> head = new ArrayList<String>();
		XSSFRow  Row=null;
		XSSFCell cell=null;
		
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
	    
//	    // TO get the destination folder to save screenshots
//		System.out.println("Enter The FILE LOCATION to save the DESTINATION FILES");
//		System.out.println("COPY AND PASTE TO AVOID ERRORS");
//		String destinationlocation = s.next();
//		System.out.println();
		
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
				
				// It is used to delay the operation					
				WebDriverWait wait = new WebDriverWait(driver, 60); 
				wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("/html/body/div[1]/div/div[3]/div/div[2]/div[1]/div[2]/table[1]/tbody/tr[2]/td"))); 
				
				
				//For Row One Titles
				if(lineNo==0){
					
					head.add("Register Number");
					head.add("Name");
					if(driver.findElements(By.xpath("//table[@id='results_subject_table']/tbody/tr[2]/td[2]")).size() != 0)
						head.add(driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[2]/td[2]")).getText());
					if(driver.findElements(By.xpath("//table[@id='results_subject_table']/tbody/tr[3]/td[2]")).size() != 0)
						head.add(driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[3]/td[2]")).getText());
					if(driver.findElements(By.xpath("//table[@id='results_subject_table']/tbody/tr[4]/td[2]")).size() != 0)
						head.add(driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[4]/td[2]")).getText());
					if(driver.findElements(By.xpath("//table[@id='results_subject_table']/tbody/tr[5]/td[2]")).size() != 0)
						head.add(driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[5]/td[2]")).getText());
					if(driver.findElements(By.xpath("//table[@id='results_subject_table']/tbody/tr[6]/td[2]")).size() != 0)
						head.add(driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[6]/td[2]")).getText());
					if(driver.findElements(By.xpath("//table[@id='results_subject_table']/tbody/tr[7]/td[2]")).size() != 0)
						head.add(driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[7]/td[2]")).getText());
					if(driver.findElements(By.xpath("//table[@id='results_subject_table']/tbody/tr[8]/td[2]")).size() != 0)
						head.add(driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[8]/td[2]")).getText());
					if(driver.findElements(By.xpath("//table[@id='results_subject_table']/tbody/tr[9]/td[2]")).size() != 0)
						head.add(driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[9]/td[2]")).getText());
					if(driver.findElements(By.xpath("//table[@id='results_subject_table']/tbody/tr[10]/td[2]")).size() != 0)
						head.add(driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[10]/td[2]")).getText());
					if(driver.findElements(By.xpath("//table[@id='results_subject_table']/tbody/tr[11]/td[2]")).size() != 0)
						head.add(driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[11]/td[2]")).getText());
					if(driver.findElements(By.xpath("//table[@id='results_subject_table']/tbody/tr[12]/td[2]")).size() != 0)
						head.add(driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[12]/td[2]")).getText());

		        	
		        	Row  = sheet.createRow(lineNo);
		        	for (int i = 0; i < head.size(); i++) 
					{	
		        		cell = Row.createCell(i);
		        		cell.setCellValue(head.get(i));
		        		sheet.autoSizeColumn(i);
					}
		        	
		        	head.clear();
		        	
		        	lineNo++;
				}
				
				// To print respective grades
				head.add(line.toString());
				head.add(driver.findElement(By.xpath("//*[@id=\"student_info\"]/tbody/tr[3]/td")).getText().substring(21));
				if(driver.findElements(By.xpath("//table[@id='results_subject_table']/tbody/tr[2]/td[7]")).size() != 0)
					head.add(driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[2]/td[7]/div")).getText());
				if(driver.findElements(By.xpath("//table[@id='results_subject_table']/tbody/tr[3]/td[7]")).size() != 0)
					head.add(driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[3]/td[7]/div")).getText());
				if(driver.findElements(By.xpath("//table[@id='results_subject_table']/tbody/tr[4]/td[7]")).size() != 0)
					head.add(driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[4]/td[7]")).getText());
				if(driver.findElements(By.xpath("//table[@id='results_subject_table']/tbody/tr[5]/td[7]")).size() != 0)
					head.add(driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[5]/td[7]")).getText());
				if(driver.findElements(By.xpath("//table[@id='results_subject_table']/tbody/tr[6]/td[7]")).size() != 0)
					head.add(driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[6]/td[7]")).getText());
				if(driver.findElements(By.xpath("//table[@id='results_subject_table']/tbody/tr[7]/td[7]")).size() != 0)
					head.add(driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[7]/td[7]")).getText());
				if(driver.findElements(By.xpath("//table[@id='results_subject_table']/tbody/tr[8]/td[7]")).size() != 0)
					head.add(driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[8]/td[7]")).getText());
				if(driver.findElements(By.xpath("//table[@id='results_subject_table']/tbody/tr[9]/td[7]")).size() != 0)
					head.add(driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[9]/td[7]")).getText());
				if(driver.findElements(By.xpath("//table[@id='results_subject_table']/tbody/tr[10]/td[7]")).size() != 0)
					head.add(driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[10]/td[7]")).getText());
				if(driver.findElements(By.xpath("//table[@id='results_subject_table']/tbody/tr[11]/td[7]")).size() != 0)
					head.add(driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[11]/td[7]")).getText());
				if(driver.findElements(By.xpath("//table[@id='results_subject_table']/tbody/tr[12]/td[7]")).size() != 0)
					head.add(driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[12]/td[7]")).getText());

				Row  = sheet.createRow(lineNo);
				
				for (int j=0; j < head.size(); j++) 
				{	
		    		cell = Row.createCell(j);
		    		cell.setCellValue(head.get(j));
		    		sheet.autoSizeColumn(j);
		    		
		    		//For console output	
//		    		System.out.println("in for");
//		    		System.out.print(head.get(j)+" ");
				}
				
				head.clear();
				lineNo++;
				
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
		workbook.write(fileOut);
		reader.close();
		workbook.close();
		fileOut.close();
		driver.quit();
		System.out.println("The Excel file created");
		


	}
}
