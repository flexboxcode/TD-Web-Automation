import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.Keys;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;

public class Samplecount {
    public static void main(String[] args) throws IOException, InterruptedException {
        // Set the path to the geckodriver executable
        String edgeDriverPath = "C:\\Selenium\\Drivers\\msedgedriver.exe"; // replace with your EdgeDriver path

        System.setProperty("webdriver.edge.driver", edgeDriverPath);
        WebDriver driver = new EdgeDriver();

        driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));

        // Initialize the start date
        LocalDate startDate = LocalDate.of(2023, 5, 1); // replace with your start date

        // Format it as per your needs
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("ddMMyyyy");

        // Navigate to the website
        driver.get("https://tdiagnostics.telangana.gov.in");

        // Enter the username
        WebElement username = driver.findElement(By.id("UserName"));
        username.sendKeys("nvhlab10");

        // Enter the password
        WebElement password = driver.findElement(By.id("Password"));
        password.sendKeys("abc1234");

        // Click the login button
        WebElement loginButton = driver.findElement(By.id("LoginButton"));
        loginButton.click();

        // Define the number of tabs
        int numberOfTabs = 31;

        // Create a new workbook
        Workbook workbook = new XSSFWorkbook();

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

            // Format the current date
            String formattedDate = startDate.format(formatter);

            // Simulate pressing the Backspace key 10 times and set the date for each date field
            for (WebElement dateField : new WebElement[]{fromDateField, toDateField}) {
                for (int i = 0; i < 12; i++) {
                    dateField.sendKeys(Keys.BACK_SPACE);
                }
                dateField.sendKeys(formattedDate);
            }

            // Increment the date for the next tab
            startDate = startDate.plusDays(1);

            // Locate the button and click it
            WebElement button = driver.findElement(By.id("ctl00_cpMiddleContent_radPnl_BillingChargeITems_i0_i0_btn_Search"));
            button.click();

            // Wait for the page to load
            Thread.sleep(2000); // wait for 2 seconds

            // Get table
            WebElement table = driver.findElement(By.id("ctl00_cpMiddleContent_radGrid_SampleCollection_ctl00"));

            // Get all rows
            List<WebElement> rows = table.findElements(By.tagName("tr"));

            // Create a new sheet for the current tab
            Sheet sheet = workbook.createSheet("May " + tabIndex );

            int rowNum = 0;

            // Loop to iterate over the rows
            for (int i = 1; i < rows.size(); i++) { // Skip the header row
                // Get columns/cells
                List<WebElement> cols = rows.get(i).findElements(By.tagName("td"));
                String facilityName = cols.get(1).getText(); // Facility Name is the second column
                String samplesCollected = cols.get(6).getText(); // Samples Collected is the seventh column

                // Create a new row in the Excel sheet
                Row excelRow = sheet.createRow(rowNum++);

                // Write the facility name and samples collected into the Excel file
                Cell facilityCell = excelRow.createCell(0);
                facilityCell.setCellValue(facilityName);
                Cell samplesCell = excelRow.createCell(1);
                samplesCell.setCellValue(samplesCollected);
            }
        }

        // Resize the columns to fit the content
        for (int i = 0; i < numberOfTabs; i++) {
            workbook.getSheetAt(i).autoSizeColumn(0);
            workbook.getSheetAt(i).autoSizeColumn(1);
        }

        // Write the workbook back to the same file
        try (FileOutputStream fileOut = new FileOutputStream("C:\\Selenium\\Excel\\TABLEData.xlsx")) {
            workbook.write(fileOut);
        }

        // Close the workbook
        workbook.close();

        
      
    }
}
