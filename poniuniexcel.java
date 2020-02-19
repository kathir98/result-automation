package pondiuniresult;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class poniuniexcel {
	public static void main(String a[]) throws FileNotFoundException, IOException
	{
	System.setProperty("webdriver.chrome.driver","F:\\selenium\\chromedriver.exe");
	WebDriver driver = new ChromeDriver();
	driver.manage().window().maximize();
	driver.get("http://exam.pondiuni.edu.in/oresults/");
	
	File excelFile=new File("F:\\Eclipseworkspaces\\resultautomation\\registernumber.xlsx");	//Path of Input File having register number in first column
	InputStream fin=new FileInputStream(excelFile);
	XSSFWorkbook wb = new XSSFWorkbook(fin);
	XSSFSheet sh = wb.getSheetAt(0);
	int rowLastNo=sh.getPhysicalNumberOfRows();
	
//	BufferedReader reader = null;
//	reader = new BufferedReader(new FileReader("registerno.txt"));
//	String line = reader.readLine();
	
	XSSFWorkbook workbook = new XSSFWorkbook();
	FileOutputStream fileOut=new FileOutputStream("results.xlsx");	//Path of Output File
	XSSFSheet sheet = workbook.createSheet("Result");
	int index=1,lineNo=0;
	ArrayList<String> head = new ArrayList<String>();
	Iterator<Row> rows = sh.rowIterator();
	XSSFRow  Row=null;
	XSSFCell cell=null;
	
	//Loop to execute automation for rowLastNo times
	while(rows.hasNext()){
		
		XSSFCell rollNo=wb.getSheetAt(0).getRow(index).getCell(0);	//Getting rollNo from input excel sheet
		
		driver.findElement(By.id("reg_no")).clear();
		driver.findElement(By.id("reg_no")).sendKeys(rollNo.toString());  //Sending rollNo values to Register Number text box
		driver.findElement(By.id("exam")).sendKeys("Seventh");	//Choosing semester from drop down list box
		driver.findElement(By.xpath("//*[@id=\"print_app_form\"]")).click();	//Clicking on submit button 
		
		WebDriverWait wait = new WebDriverWait(driver, 60);		//Driver wait
		wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("/html/body/div[1]/div/div[3]/div/div[2]/div[1]/div[2]/table[1]/tbody/tr[2]/td"))); 
		
		//For Row One Titles
		if(lineNo==0){
			
			head.add("Register Number");
			head.add("Name");
        	head.add(driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[2]/td[2]")).getText());
//        	System.out.println(driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[2]/td[2]")).getText());
        	head.add(driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[3]/td[2]")).getText());
//       	System.out.println(driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[3]/td[2]")).getText());
        	head.add(driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[4]/td[2]")).getText());
//        	System.out.println(driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[4]/td[2]")).getText());
        	head.add(driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[5]/td[2]")).getText());
//        	System.out.println(driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[5]/td[2]")).getText());
        	head.add(driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[6]/td[2]")).getText());
//        	System.out.println(driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[6]/td[2]")).getText());
        	head.add(driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[7]/td[2]")).getText());
//        	System.out.println(driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[7]/td[2]")).getText());
        	head.add(driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[8]/td[2]")).getText());
//        	System.out.println(driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[8]/td[2]")).getText());
        	head.add(driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[9]/td[2]")).getText());
//        	System.out.println(driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[9]/td[2]")).getText());
//        	head.add(driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[10]/td[2]")).getText());
//        	System.out.println(driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[10]/td[2]")).getText());
//        	head.add(driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[11]/td[2]")).getText());
//        	System.out.println(driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[11]/td[2]")).getText());
        	
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
		// To print the respective grades
		head.add(rollNo.toString());
		head.add(driver.findElement(By.xpath("//*[@id=\"student_info\"]/tbody/tr[3]/td")).getText().substring(21));
		head.add(driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[2]/td[7]/div")).getText());
		head.add(driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[3]/td[7]/div")).getText());
		head.add(driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[4]/td[7]")).getText());
		head.add(driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[5]/td[7]")).getText());
		head.add(driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[6]/td[7]")).getText());
		head.add(driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[7]/td[7]")).getText());
		head.add(driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[8]/td[7]")).getText());
		head.add(driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[9]/td[7]")).getText());
//		head.add(driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[10]/td[7]")).getText());
//		head.add(driver.findElement(By.xpath("//table[@id='results_subject_table']/tbody/tr[11]/td[7]")).getText());
		
		Row  = sheet.createRow(lineNo);
		
		for (int j=0; j < head.size(); j++) 
		{	
    		cell = Row.createCell(j);
    		cell.setCellValue(head.get(j));
    		sheet.autoSizeColumn(j);
    		
    		//For console output	
//    		System.out.println("in for");
//    		System.out.print(head.get(j)+" ");
		}
		System.out.println();
		head.clear();
		lineNo++;
    	index++;
    	driver.navigate().back();
		if(lineNo==(rowLastNo+1)){
			workbook.write(fileOut);
			break;
		}
	}
	fin.close();
	wb.close();
	workbook.close();
	fileOut.close();
	driver.quit();

}
}
