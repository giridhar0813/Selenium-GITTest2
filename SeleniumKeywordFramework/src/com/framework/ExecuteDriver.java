package com.framework;

import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.formula.functions.Rows;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Parameters;
import org.testng.annotations.Test;

import com.aventstack.extentreports.Status;
import com.aventstack.extentreports.markuputils.ExtentColor;
import com.aventstack.extentreports.markuputils.MarkupHelper;

public class ExecuteDriver extends BaseTest {
	WebDriver driver;
	String keyword ;
	String id ;
	String locator;
	String locatervalue ;
	String testData ;
	int distAppCount=0;
	String ISV;
	String isvStatus;
	boolean status;




	//public ExecuteDriver() throws Exception{

	//	readDataFromSuiteExcel(Constants.Master_Data_Sheet,"TestSuite");


	//}


	@Test
	public void executeTests() throws IOException, InterruptedException{

		readDataFromSuiteExcel(Constants.Master_Data_Sheet,"TestSuite");
		System.out.println(totalRows);
		
		//System.out.println("test123:");
		
		//System.out.println("test123:");

			System.out.println("changes pushed from workpace 3");

		for(int i =1;i<=totalRows;i++){

			System.out.println("value of i is "+ i);
			if( s.getRow(i).getCell(0) != null){

				ISV = s.getRow(i).getCell(0).getStringCellValue();
				isvStatus = s.getRow(i).getCell(2).getStringCellValue();
				distAppCount = (int) s.getRow(i).getCell(3).getNumericCellValue();
			}

			if(isvStatus.equals("Yes"))
			{
				if( (s.getRow(i).getCell(2) != null) && (distAppCount < 1)){


					String runStatus = s.getRow(i).getCell(2).getStringCellValue();

					System.out.println(runStatus);
					String distributedApp = s.getRow(i).getCell(1).getStringCellValue();
					logger = reports.createTest(ISV  + ": " + distributedApp);

					if(runStatus.equals("Yes")){		


						//String distributedApp = s.getRow(i).getCell(1).getStringCellValue();
						System.out.println( distributedApp);
						//logger = reports.createTest(ISV  + ": " + distributedApp);

						readDataFromTestDatal(Constants.Master_Data_Sheet,distributedApp);
						System.out.println(totalRows);

						loop:	for(int j=1;j<=totalTestRows;j++){

							keyword=null;
							locator=null;
							locatervalue=null;
							testData=null;

							if( (tSheet.getRow(j).getCell(0) != null))					    
								keyword = tSheet.getRow(j).getCell(0).getStringCellValue();
							if( (tSheet.getRow(j).getCell(1) != null))
								locator = tSheet.getRow(j).getCell(1).getStringCellValue();
							if( (tSheet.getRow(j).getCell(2) != null))
								locatervalue = tSheet.getRow(j).getCell(2).getStringCellValue();
							if( (tSheet.getRow(j).getCell(3) != null))
								testData = tSheet.getRow(j).getCell(3).getStringCellValue();

							System.out.println(keyword + " " + locator + " " + locatervalue + " " + testData);


							switch(keyword){

							case "NavigateURL":

								status=executeAction(keyword,locator,locatervalue,testData);
								if(status== false){

									logger.log(Status.INFO, MarkupHelper.createLabel("Skipping the rest of the steps",ExtentColor.BLACK));
									writeResultToExcel(Constants.Master_Data_Sheet,"Failed",i);
									break loop;

								}
								else
									System.out.println("inside true");
									writeResultToExcel(Constants.Master_Data_Sheet,"Success",i);

								break;

							case "enterText":

								status=executeAction(keyword,locator,locatervalue,testData);
								if(status== false){

									logger.log(Status.INFO, MarkupHelper.createLabel("Skipping the rest of the steps",ExtentColor.BLACK));
									writeResultToExcel(Constants.Master_Data_Sheet,"Failed",i);
									break loop;

								}
								else
									System.out.println("inside true");
									writeResultToExcel(Constants.Master_Data_Sheet,"Success",i);

								break;

							case "clickElement":

								status=executeAction(keyword,locator,locatervalue,testData);
								if(status== false){

									logger.log(Status.INFO, MarkupHelper.createLabel("Skipping the rest of the steps",ExtentColor.BLACK));
									writeResultToExcel(Constants.Master_Data_Sheet,"Failed",i);
									break loop;

								}
								else
									System.out.println("inside true");
									writeResultToExcel(Constants.Master_Data_Sheet,"Success",i);
									
								break;

							default:

								break;



							}

						}

					}

					else
						logger.log(Status.SKIP, MarkupHelper.createLabel(distributedApp + " exection is skipped", ExtentColor.YELLOW));



				}
				else
					distAppCount=0;
			}

			else
				i+=distAppCount;


		}




		w.close();


	}

















}
