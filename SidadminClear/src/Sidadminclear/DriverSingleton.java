package Sidadminclear;
import org.openqa.selenium.WebDriver;

import org.openqa.selenium.edge.EdgeDriver;

import java.time.Duration;


public class DriverSingleton {

    private static WebDriver driver;


    public static WebDriver getDriver() {

        if (driver == null || !isDriverActive()) {

            driver = new EdgeDriver(); // or FirefoxDriver(), etc.

            driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));

        }

        return driver;

    }


    private static boolean isDriverActive() {

        try {

            driver.getTitle();

            return true;

        } catch (Exception e) {

            return false;

        }

    }

}