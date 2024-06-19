package tests;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.Iterator;

public class UploadDownloadExample
{



    //Get the column number of the Price Column
    //Get the row number of the price of Apple row

    public static void main(String[] args) throws IOException, InterruptedException {


        WebDriver driver = new ChromeDriver();
        driver.manage().window().maximize();
        driver.manage().deleteAllCookies();
        driver.get("https://rahulshettyacademy.com/upload-download-test/index.html");
        driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(5));
        WebElement downloadButton = driver.findElement(By.xpath("//button[@id='downloadButton']"));
        downloadButton.click();
        //driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(5));

        Thread.sleep(5000);

        String fruitName= "Apple";
         String fileName = System.getProperty("user.home")+"\\Downloads\\download.xlsx";
        driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(5));

        File file =  new File(fileName);
        System.out.println(file.exists());




        int col = getColumnNumber(fileName,"price");
        int row = getRowNumber(fileName,"Apple");
        Assert.assertTrue(updateCell(fileName,row,col,"600"));
        WebElement uploadButton = driver.findElement(By.xpath("//input[@id='fileinput']"));
        uploadButton.sendKeys(fileName);
        WebDriverWait wait = new WebDriverWait(driver,Duration.ofSeconds(5));
        By validationMsg = By.xpath("//div[@class = 'Toastify__toast-body']/div[2]");
        wait.until(ExpectedConditions.visibilityOfElementLocated(validationMsg));
        String actualValidationMsg = driver.findElement(validationMsg).getText();
        System.out.println(actualValidationMsg);
        String expectedMsg = "Updated Excel Data Successfully.";
        Assert.assertEquals(actualValidationMsg,expectedMsg);
        wait.until(ExpectedConditions.invisibilityOfElementLocated(validationMsg));
        driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(5));
        String priceColumn = driver.findElement(By.xpath("//div[text()='Price']")).getAttribute("data-column-id");
        String actualPrice = driver.findElement(By.xpath("//div[text()='"+fruitName+"' ]/parent::div/parent::div/div[@id='cell-"+priceColumn+"-undefined']")).getText();
        Assert.assertEquals(600,actualPrice);
        driver.close();


    }
    private static int getColumnNumber(String fileName, String colName) throws IOException {
        FileInputStream fis = new FileInputStream(fileName);
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        XSSFSheet sheet = workbook.getSheet("Sheet1");
        Iterator<Row> rows = sheet.iterator();
        Row firstRow = rows.next();
        Iterator<Cell> ce = firstRow.cellIterator();
        int k = 1;
        int column = 0;

        while (ce.hasNext())
        {
            Cell value = ce.next();
            if (value.getStringCellValue().equalsIgnoreCase(colName))
            {
                column = k;

            }
            k++;
        }
        System.out.println(column);


        return column;
    }



    private static int getRowNumber(String fileName, String rowText) throws IOException {
        FileInputStream fis = new FileInputStream(fileName);
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        XSSFSheet sheet = workbook.getSheet("Sheet1");
        Iterator<Row> rows = sheet.iterator();
        int k = 1;
        int rowIndex = -1;
        while (rows.hasNext()) {
            Row row = rows.next();
            Iterator<Cell> cells = row.cellIterator();
            while (cells.hasNext())
            {
                Cell cell = cells.next();
                if (cell.getCellType()== CellType.STRING && cell.getStringCellValue().equalsIgnoreCase(rowText))
                {
                    rowIndex= k;
                }
            }
            k++;
        }



        return rowIndex;
    }

    private static boolean updateCell(String fileName, int row, int col, String number) throws IOException {
        FileInputStream fis = new FileInputStream(fileName);
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        XSSFSheet sheet = workbook.getSheet("Sheet1");
       Row rowField =  sheet.getRow(row-1);
       Cell cellField = rowField.getCell(col-1);
       cellField.setCellValue(number);
        FileOutputStream fos = new FileOutputStream(fileName);
        workbook.write(fos);
        workbook.close();
        fis.close();
        return true;


    }


}
