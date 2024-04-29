package Sidadminclear;
import java.io.FileInputStream;

import java.io.IOException;

import java.time.Duration;


import org.openqa.selenium.By;

import org.openqa.selenium.WebDriver;

import org.openqa.selenium.WebElement;

import org.openqa.selenium.edge.EdgeDriver;

import org.openqa.selenium.support.ui.ExpectedConditions;

import org.openqa.selenium.support.ui.WebDriverWait;


import org.apache.poi.ss.usermodel.Sheet;

import org.apache.poi.ss.usermodel.Workbook;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class sidadminclear {
    public static void main(String[] args) throws IOException, InterruptedException {
        String excelFilePath = "C:\\Selenium\\Excel\\aadhar.xlsx";
        WebDriver driver = DriverSingleton.getDriver();
        driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));

        try {
            FileInputStream excelFile = new FileInputStream(excelFilePath);
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet sheet = workbook.getSheetAt(0); // Assuming data is on the first sheet

            int rowCount = sheet.getPhysicalNumberOfRows();

            driver.get("https://tdiagnostics.telangana.gov.in/LabUser/UniqueDataReportNew.aspx");
            driver.findElement(By.id("UserName")).sendKeys("sidadmin");
            driver.findElement(By.id("Password")).sendKeys("ntpl@1234");
            driver.findElement(By.id("LoginButton")).click();
            Thread.sleep(2000);

            for (int row = 0; row < rowCount; row++) {
                long numericValue = (long) sheet.getRow(row).getCell(0).getNumericCellValue();
                String aadhar = String.valueOf(numericValue);

              
                WebElement Aadhar = driver.findElement(By.id("ctl00_cpMiddleContent_txtId"));
               

                Aadhar.sendKeys(aadhar);

              
                driver.findElement(By.id("btn_Search")).click();
               
                driver.findElement(By.id("edit_button_1")).click();
           

                WebElement selectaadhar = driver.findElement(By.id("SerachId_text1"));

                selectaadhar.click();

                selectaadhar.clear();

          

                driver.findElement(By.id("save_button_1")).click();

                // Close the current window or tab
                driver.close();

                // Open a new instance of the browser
                driver = DriverSingleton.getDriver();
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            // Close the browser session
            // driver.quit();
        }
    }
}