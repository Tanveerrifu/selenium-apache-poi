import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;

public class GoogleSearchScraper {
    public static void main(String[] args) {
        String inputFile = "C:\\Users\\tanvi\\Documents\\ExcelSearch.xlsx";
        String outputFile = "C:\\Users\\tanvi\\Documents\\ExcelSearch2.xlsx";
        try {

            List<String> keywords = readKeywordsFromExcel(inputFile);


            List<String> results = performGoogleSearch(keywords);


            String shortestKeyword = findShortestKeyword(keywords);
            String longestKeyword = findLongestKeyword(keywords);

            saveResultsToExcel(outputFile, results, shortestKeyword, longestKeyword);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static List<String> readKeywordsFromExcel(String inputFile) throws IOException {
        List<String> keywords = new ArrayList<>();
        FileInputStream fis = new FileInputStream(inputFile);
        Workbook workbook = WorkbookFactory.create(fis);
        Sheet sheet = workbook.getSheetAt(0);

        for (Row row : sheet) {
            Cell cell = row.getCell(0);
            if (cell != null && cell.getCellType() == CellType.STRING) {
                keywords.add(cell.getStringCellValue());
            }
        }

        workbook.close();
        fis.close();
        return keywords;
    }

    private static List<String> performGoogleSearch(List<String> keywords) {
        System.setProperty("webdriver.chrome.driver", "G:\\QA CODE\\chromedriver_win32\\chromedriver.exe");
        WebDriver driver = new ChromeDriver();

        List<String> results = new ArrayList<>();
        for (String keyword : keywords) {
            driver.get("https://www.google.com");

            WebElement searchBox = driver.findElement(By.name("q"));
            Actions actions = new Actions(driver);
            actions.sendKeys(searchBox, keyword).sendKeys(Keys.ENTER).build().perform();

            driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);

            WebElement searchResult = driver.findElement(By.cssSelector("h3"));
            results.add(searchResult.getText());
        }

        driver.quit();
        return results;
    }

    private static String findShortestKeyword(List<String> keywords) {
        return Collections.min(keywords, Comparator.comparingInt(String::length));
    }

    private static String findLongestKeyword(List<String> keywords) {
        return Collections.max(keywords, Comparator.comparingInt(String::length));
    }

    private static void saveResultsToExcel(String outputFile, List<String> results, String shortestKeyword, String longestKeyword) throws IOException {
        Workbook workbook = WorkbookFactory.create(true);
        Sheet sheet = workbook.createSheet();


        for (int i = 0; i < results.size(); i++) {
            Row row = sheet.createRow(i);
            Cell cell = row.createCell(0);
            cell.setCellValue(results.get(i));
        }


        Row shortestRow = sheet.createRow(results.size());
        Cell shortestCell = shortestRow.createCell(0);
        shortestCell.setCellValue("Shortest Keyword:");
        Cell shortestValueCell = shortestRow.createCell(1);
        shortestValueCell.setCellValue(shortestKeyword);


        Row longestRow = sheet.createRow(results.size() + 1);
        Cell longestCell = longestRow.createCell(0);
        longestCell.setCellValue("Longest Keyword:");
        Cell longestValueCell = longestRow.createCell(1);
        longestValueCell.setCellValue(longestKeyword);

        FileOutputStream fos = new FileOutputStream(outputFile);
        workbook.write(fos);
        workbook.close();
        fos.close();
    }
}
