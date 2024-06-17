package tests;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class DataProviderClass
{
    DataFormatter formatter = new DataFormatter();
    @Test (dataProvider = "driveTest")
    public void firsTest(String brand,String model,String variant, String value) throws IOException {
                   //getData();

        System.out.println("The"+" " +brand + " " +model + " " +variant+" "+"variant is" +value);

    }

    @DataProvider (name = "driveTest")

    public Object[][] getData() throws IOException {

        FileInputStream fis = new FileInputStream(System.getProperty("user.dir")+"//src//test//java//data//dataSheet.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        XSSFSheet sheet = workbook.getSheetAt(0);
        int rowCount = sheet.getPhysicalNumberOfRows();
        XSSFRow rows = sheet.getRow(0);
        int columnCount = rows.getLastCellNum();
        Object dataObj[][] = new Object[rowCount-1][columnCount];
        for (int i = 0; i<rowCount-1;i++)
        {
            rows = sheet.getRow(i+1);
            for (int j = 0; j<columnCount;j++)
            {
                XSSFCell cell = rows.getCell(j);
                dataObj[i][j] = formatter.formatCellValue(cell);

            }
        }

return dataObj;
    }





}
