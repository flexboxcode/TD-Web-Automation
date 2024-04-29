package gadadmin;

import org.openqa.selenium.By;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.Keys;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.ArrayList;
import java.util.List;
import java.io.File;


import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;

import java.time.Duration;
import java.util.ArrayList;
import java.util.List;

public class gadadmin {
    public static void main(String[] args) throws IOException {
        // Set the path to the geckodriver executable
    	  String edgeDriverPath = "C:\\Selenium\\Drivers\\msedgedriver.exe"; // replace with your EdgeDriver path

          System.setProperty("webdriver.edge.driver", edgeDriverPath);
          WebDriver driver = new EdgeDriver();

        driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));
     // Get yesterday's date
        LocalDate yesterday = LocalDate.now().minusDays(1);

        // Format it as per your needs
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("ddMMyyyy");
        String formattedDate = yesterday.format(formatter);


        // Navigate to the website
        driver.get("https://tdiagnostics.telangana.gov.in");

        // Enter the username
        WebElement username = driver.findElement(By.id("UserName"));
        username.sendKeys("gadadmin");

        // Enter the password
        WebElement password = driver.findElement(By.id("Password"));
        password.sendKeys("ntpl@1234");

        // Click the login button
        WebElement loginButton = driver.findElement(By.id("LoginButton"));
        loginButton.click();
        
     // Define the number of tabs
        int numberOfTabs = 13;


     // Create a new workbook and sheet
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Footer Data");

     int rowNum = 0; // Initialize row number outside the loop

  

     for (int tabIndex = 1; tabIndex <= numberOfTabs; tabIndex++) {
         // Open a new tab
         ((JavascriptExecutor) driver).executeScript("window.open();");

         // Switch to the new tab
         ArrayList<String> tabs = new ArrayList<>(driver.getWindowHandles());
         driver.switchTo().window(tabs.get(tabIndex)); // switch to the new tab

         // Navigate to the web page
         driver.get("https://tdiagnostics.telangana.gov.in/Reporting/SampleCollectionBasedOnFacilityNew.aspx");

         // Locate the date field elements
         WebElement fromDateField = driver.findElement(By.id("ctl00_cpMiddleContent_radPnl_BillingChargeITems_i0_i0_radMaskedtxt_FromDate_text"));
         WebElement toDateField = driver.findElement(By.id("ctl00_cpMiddleContent_radPnl_BillingChargeITems_i0_i0_radMaskedtxt_ToDate_text"));

         // Simulate pressing the Backspace key 10 times and set the date for each date field
         for (WebElement dateField : new WebElement[]{fromDateField, toDateField}) {
             for (int i = 0; i < 12; i++) {
                 dateField.sendKeys(Keys.BACK_SPACE);
             }
             dateField.sendKeys(formattedDate);
         }

         // Locate the dropdown field element
         WebElement dropdownField = driver.findElement(By.id("ctl00_cpMiddleContent_radPnl_BillingChargeITems_i0_i0_radCmb_Hub_Input"));

         // Click the down arrow key a certain number of times
         for (int i = 0; i < tabIndex; i++) {
             dropdownField.sendKeys(Keys.ARROW_DOWN);
         }

         // Locate the button and click it
         WebElement button = driver.findElement(By.id("ctl00_cpMiddleContent_radPnl_BillingChargeITems_i0_i0_btn_Search"));
         button.click();

         // Locate the footer cells
         List<WebElement> footerCells = driver.findElements(By.cssSelector(".rgFooter > td"));

         // Write the text of each cell into the Excel file
         Row row = sheet.createRow(rowNum++); // Create a new row for each tab

         int cellNum = 0;
         for (WebElement cell : footerCells) {
             Cell excelCell = row.createCell(cellNum++);
             excelCell.setCellValue(cell.getText());
         }
     }

     // Resize the columns to fit the content
     for (int i = 0; i < numberOfTabs; i++) {
         sheet.autoSizeColumn(i);
     }

     // Write the workbook back to the same file
     try (FileOutputStream fileOut = new FileOutputStream("C:\\Selenium\\Excel\\GadadminData.xlsx")) {
         workbook.write(fileOut);
     }

     // Close the workbook and FileInputStream
     workbook.close();
   
   
    }
}

