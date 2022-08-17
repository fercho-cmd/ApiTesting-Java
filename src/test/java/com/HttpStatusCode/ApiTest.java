package com.HttpStatusCode;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import org.openqa.selenium.WebDriver;
import org.testng.annotations.AfterTest;
import org.testng.annotations.Test;

import com.HttpStatusCode.ExcelUtilities;

import io.restassured.RestAssured;
import jxl.Workbook;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

public class ApiTest {

	WebDriver driver;
	public static WritableWorkbook workbook;
	

	@Test
	public void PostApiTest() throws Exception {
		
		

		ExcelUtilities eu = new ExcelUtilities();

		String inputPath = "./InputFolder";
		ArrayList<String>  title = eu.getColumnDataAsList(inputPath, "/PostApiTest.xlsx", "Methods", "Title");
		ArrayList<String>  urls = eu.getColumnDataAsList(inputPath, "/PostApiTest.xlsx", "Methods", "Urls");
		ArrayList<String>  body = eu.getColumnDataAsList(inputPath, "/PostApiTest.xlsx", "Methods", "Body");
		ArrayList<String>  response = eu.getColumnDataAsList(inputPath, "/PostApiTest.xlsx", "Methods", "Expected Response");
		ArrayList<String>  statuscode = eu.getColumnDataAsList(inputPath, "/PostApiTest.xlsx", "Methods", "Expected Status Code");


		// Create Output file for Report
		String report = ".//report//PostResults.xls";
		File outputfile = new File(report);// output file
		workbook = Workbook.createWorkbook(outputfile);
		// Format and Style
		WritableCellFormat cFormat = new WritableCellFormat();
		WritableFont font = new WritableFont(WritableFont.ARIAL, 14, WritableFont.BOLD);
		cFormat.setFont(font);

		// Create Results File
		WritableSheet wSheet = workbook.createSheet("Results", 1);
		// adding column headers to the output results file
		wSheet.addCell(new jxl.write.Label(0, 0, "Title", cFormat));
		wSheet.addCell(new jxl.write.Label(1, 0, "Url", cFormat));
		wSheet.addCell(new jxl.write.Label(2, 0, "Body", cFormat));
        wSheet.addCell(new jxl.write.Label(3, 0, "Actual Status Code", cFormat));
        wSheet.addCell(new jxl.write.Label(4, 0, "Expected Status Code", cFormat));
		wSheet.addCell(new jxl.write.Label(5, 0, "Actual Response", cFormat));
		wSheet.addCell(new jxl.write.Label(6, 0, "Expected Response", cFormat));

        
		
         
		 //Set Index start 	for URL ArrayList
		    int index = 0;
		    
		 // Take the Body per row
			for (String testbody : body) {
				
				//Return the URL using Index Position from Url ArrayList
				ArrayList<String> list = new ArrayList<String>(urls);
				String url = list.get(index);
				ArrayList<String> Title = new ArrayList<String>(title);
				String titles = Title.get(index);
				ArrayList<String> Response = new ArrayList<String>(response);
				String expectedResponse = Response.get(index);
				ArrayList<String> Code = new ArrayList<String>(statuscode);
				String expectedCode = Code.get(index);
				
				++ index;

		//	String url = "https://api.qawyre.com/v3/transfers";

				
				// Login and Extract Session Id from Response
				String SessionId = RestAssured.given().contentType("application/json")
						.body("{\"password\":\"Excel2021*1\",\"email\":\"Fernando@sendwyre.com\"}").when()
						.post("https://admin.qawyre.com/ajax/core/v2/sessions/auth").then().assertThat().extract()
						.path("sessionId");
				
				// Login and Extract API Key
				String secretKey = RestAssured.given().contentType("application/json").auth().oauth2(SessionId).and()
						.body("{\n"+ "\"type\": \"FULL\"\n"+ "}").when()
						.post("https://dash.qawyre.com/core/v2/apiKeys").then().assertThat().extract()
						.path("secretKey");
				
				System.out.println(SessionId);

				// TransferPostMethod
				String ResponseBody = RestAssured.given().contentType("application/json").auth()
						.oauth2(secretKey).and()
						.body(testbody).when().post(url)
						.getBody().asString();

				// Get Status Code
				int code = RestAssured.given().contentType("application/json").auth().oauth2(secretKey).and().body(testbody)
						.when().post(url).statusCode();
				String actualCodeStatus = "'"+code + "'";
				

				int j = wSheet.getRows();
				// add the results to the output excel
				wSheet.addCell(new jxl.write.Label(0, j, titles));
				wSheet.addCell(new jxl.write.Label(1, j, url));
				wSheet.addCell(new jxl.write.Label(2, j, testbody));
				wSheet.addCell(new jxl.write.Label(3, j, actualCodeStatus));
				wSheet.addCell(new jxl.write.Label(4, j, expectedCode));
				wSheet.addCell(new jxl.write.Label(5, j, ResponseBody));
				wSheet.addCell(new jxl.write.Label(6, j, expectedResponse));
				


			}

		}
	
	
	
	@Test
	public void GetApiTest() throws Exception {
		
		

		ExcelUtilities eu = new ExcelUtilities();

		String inputPath = "./InputFolder";
		ArrayList<String>  title = eu.getColumnDataAsList(inputPath, "/GetApiTest.xlsx", "Methods", "Title");
		ArrayList<String>  urls = eu.getColumnDataAsList(inputPath, "/GetApiTest.xlsx", "Methods", "Urls");
		ArrayList<String>  response = eu.getColumnDataAsList(inputPath, "/GetApiTest.xlsx", "Methods", "Expected Response");
		ArrayList<String>  statuscode = eu.getColumnDataAsList(inputPath, "/GetApiTest.xlsx", "Methods", "Expected Status Code");


		// Create Output file for Report
		String report = ".//report//GetResults.xls";
		File outputfile = new File(report);// output file
		workbook = Workbook.createWorkbook(outputfile);
		// Format and Style
		WritableCellFormat cFormat = new WritableCellFormat();
		WritableFont font = new WritableFont(WritableFont.ARIAL, 14, WritableFont.BOLD);
		cFormat.setFont(font);

		// Create Results File
		WritableSheet wSheet = workbook.createSheet("Results", 1);
		// adding column headers to the output results file
		wSheet.addCell(new jxl.write.Label(0, 0, "Title", cFormat));
		wSheet.addCell(new jxl.write.Label(1, 0, "Url", cFormat));
        wSheet.addCell(new jxl.write.Label(2, 0, "Actual Status Code", cFormat));
        wSheet.addCell(new jxl.write.Label(3, 0, "Expected Status Code", cFormat));
		wSheet.addCell(new jxl.write.Label(4, 0, "Actual Response", cFormat));
		wSheet.addCell(new jxl.write.Label(5, 0, "Expected Response", cFormat));

        
		
         
		 //Set Index start 	for URL ArrayList
		    int index = 0;
		    
		 // Take the Body per row
			for (String titles : title) {
				
				//Return the URL using Index Position from Url ArrayList
				ArrayList<String> list = new ArrayList<String>(urls);
				String url = list.get(index);
				ArrayList<String> Response = new ArrayList<String>(response);
				String expectedResponse = Response.get(index);
				ArrayList<String> Code = new ArrayList<String>(statuscode);
				String expectedCode = Code.get(index);
				
				++ index;

		//	String url = "https://api.qawyre.com/v3/transfers";

				
				// Login and Extract Session Id from Response
				String SessionId = RestAssured.given().contentType("application/json")
						.body("{\"password\":\"Kc654718\",\"email\":\"Olha@sendwyre.com\"}").when()
						.post("https://admin.qawyre.com/ajax/core/v2/sessions/auth").then().assertThat().extract()
						.path("sessionId");
				
				System.out.println(SessionId);

				
				// Login and Extract API Key
				String secretKey = RestAssured.given().contentType("application/json").auth().oauth2(SessionId).and()
						.body("{\n"+ "\"type\": \"FULL\"\n"+ "}").when()
						.post("https://dash.qawyre.com/core/v2/apiKeys").then().assertThat().extract()
						.path("secretKey");

				// TransferPostMethod
				String ResponseBody = RestAssured.given().contentType("application/json").auth()
						.oauth2(secretKey).and()
						.when().get(url)
						.getBody().asString();

				// Get Status Code
				int code = RestAssured.given().contentType("application/json").auth()
						.oauth2(secretKey).and()
						.when().get(url).statusCode();
				String actualCodeStatus = "'"+code + "'";
				

				int j = wSheet.getRows();
				// add the results to the output excel
				wSheet.addCell(new jxl.write.Label(0, j, titles));
				wSheet.addCell(new jxl.write.Label(1, j, url));
				wSheet.addCell(new jxl.write.Label(2, j, actualCodeStatus));
				wSheet.addCell(new jxl.write.Label(3, j, expectedCode));
				wSheet.addCell(new jxl.write.Label(4, j, ResponseBody));
				wSheet.addCell(new jxl.write.Label(5, j, expectedResponse));
				


			}

		}
	
	

	@AfterTest
	public void afterTest() throws IOException, WriteException {

		workbook.write();
		workbook.close();
	}

}