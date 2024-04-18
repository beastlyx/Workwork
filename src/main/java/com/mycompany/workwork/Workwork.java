package com.mycompany.workwork;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.safari.SafariDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class Workwork {
    public static void main(String[] args) throws IOException {
        WebDriver driver = new SafariDriver();
        driver.get("https://www.powerball.com/previous-results?gc=powerball&sd=2023-01-02&ed=2023-09-06");

        JavascriptExecutor js = (JavascriptExecutor) driver;
        WebDriverWait wait = new WebDriverWait(driver, 2);

        try {
            while (true) {
                WebElement loadMoreButton = wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector("#loadMore")));
                js.executeScript("arguments[0].click();", loadMoreButton);
                Thread.sleep(500);
            }
        } catch (Exception e) {
            System.out.println("No more 'Load More' button found");
        } finally {
            List<WebElement> whiteBalls = driver.findElements(By.cssSelector(".white-balls.item-powerball"));
            List<WebElement> powerBalls = driver.findElements(By.cssSelector(".powerball.item-powerball"));

            ArrayList<String> whiteNumbers = new ArrayList<>(whiteBalls.size());
            ArrayList<String> powerballNumbers = new ArrayList<>(powerBalls.size());

            for (WebElement element : whiteBalls) {
                whiteNumbers.add(element.getText());
            }

            for (WebElement element : powerBalls) {
                powerballNumbers.add(element.getText());
            }

            saveToExcel(whiteNumbers, powerballNumbers);  // Save to Excel once at the end
            driver.quit();
        }
    }

    private static void saveToExcel(ArrayList<String> whiteNumbers, ArrayList<String> powerballNumbers) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Powerball Numbers");
        int rowNum = 0;
        int colNum = 0;
        int index = 0;
        Row row = sheet.createRow(rowNum);

        for (int i = 0; i < whiteNumbers.size(); i++) {
            if (i % 5 == 0 && i != 0) {
                Cell cell = row.createCell(colNum++);
                cell.setCellValue(powerballNumbers.get(index++));
                rowNum++;
                colNum = 0;
                row = sheet.createRow(rowNum);
            }
            Cell cell = row.createCell(colNum++);
            cell.setCellValue(whiteNumbers.get(i));
        }
        Cell cell = row.createCell(colNum);
        cell.setCellValue(powerballNumbers.get(index));

        String desktopPath = System.getProperty("user.home") + "/Desktop/";
        try (FileOutputStream fileOut = new FileOutputStream(desktopPath + "powerball_numbers.xlsx")) {
            workbook.write(fileOut);
        }
        workbook.close();
    }

}