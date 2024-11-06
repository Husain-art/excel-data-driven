package hcyclewala;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;
import org.testng.annotations.Test;

public class UploadDownload {

    @Test
    public void checking() throws IOException, InterruptedException{
        System.setProperty("webdriver.edge.driver", "C:\\Users\\v-hcyclewala\\Downloads\\edgedriver_win64 (6)\\msedgedriver.exe");
        WebDriver driver = new EdgeDriver();
        driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(60));

        driver.get("https://rahulshettyacademy.com/upload-download-test/index.html");

        driver.findElement(By.id("downloadButton")).click();
        Thread.sleep(2000);

        int UpdateValue = 188;
        int columnNo = getColumnNo("C:\\Users\\v-hcyclewala\\Downloads\\download.xlsx", "price");
        System.out.println("columnNo: " + columnNo);
        int rowNo = getRowNo("C:\\Users\\v-hcyclewala\\Downloads\\download.xlsx", "Apple");
        System.out.println("rowNo:" + rowNo);
        setcellvalue("C:\\Users\\v-hcyclewala\\Downloads\\download.xlsx", columnNo, rowNo, UpdateValue);


        WebElement upload =driver.findElement(By.id("fileinput"));
        upload.sendKeys("C:\\Users\\v-hcyclewala\\Downloads\\download.xlsx");

        // //div[@class="Toastify__toast-body"]//div[2]
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(30));
        wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div[@class='Toastify__toast-body']//div[2]")));
        String status = driver.findElement(By.xpath("//div[@class='Toastify__toast-body']//div[2]")).getText();
        System.out.println(status);
        Assert.assertEquals(status, "Updated Excel Data Successfully.");
        wait.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//div[@class='Toastify__toast-body']//div[2]")));
        System.out.println("Disappear");
        String fruit = "Apple";
        int priceColumn = 4;
        String price = driver.findElement(By.xpath("//div[text() = '" + fruit + "']/parent :: div/parent :: div/div[@id='cell-" + priceColumn + "-undefined']/div")).getText();
        System.out.println(price);
        int intPrice =Integer.parseInt(price);
        Assert.assertEquals(UpdateValue, intPrice);
    }

    private int getColumnNo(String filepath, String ColumnName) throws IOException{
        FileInputStream fis = new FileInputStream(filepath);
        XSSFWorkbook workbook = new XSSFWorkbook(fis);

        int columnNo = 0;
        int noOfSheet = workbook.getNumberOfSheets();
        for(int i=0; i<noOfSheet; i++)
        {
            if(workbook.getSheetAt(i).getSheetName().equalsIgnoreCase("Sheet1"));
            {
                XSSFSheet sheet = workbook.getSheetAt(i);

                Iterator<Row> rows = sheet.iterator();
                Row firstRow = rows.next();

                
                Iterator<Cell> cells = firstRow.iterator();
                while(cells.hasNext())
                {
                    Cell cell = cells.next();
                    if(cell.getStringCellValue().equalsIgnoreCase(ColumnName))
                    break;
                    else
                    columnNo++;

                }
            }
        }
        return columnNo;        
    }

    private int getRowNo(String fileName, String value) throws IOException{
        FileInputStream fis = new FileInputStream(fileName);
        XSSFWorkbook workbook = new XSSFWorkbook(fis);

        int noOfSheet = workbook.getNumberOfSheets();
        int k=0;
        int rowIndex=-1;
        for(int i=0; i<noOfSheet; i++)
        {
            if(workbook.getSheetAt(i).getSheetName().equalsIgnoreCase("Sheet1"));
            {
                XSSFSheet sheet = workbook.getSheetAt(i);

                Iterator<Row> rows = sheet.iterator();

                while(rows.hasNext())
                {
                    Row nextRow = rows.next();               
                    Iterator<Cell> cells = nextRow.cellIterator();
                    while(cells.hasNext())
                    {
                        Cell cell = cells.next();
                        if(cell.getCellType()==CellType.STRING && cell.getStringCellValue().equalsIgnoreCase(value))
                        {
                            rowIndex=k;
                        }
                    
                    }
                    k++;
                }
                return rowIndex; 
            }
        }
        return rowIndex;
            
    }

    private void setcellvalue(String fileName, int columnNo, int rowNo, int Value) throws IOException{
        FileInputStream fis = new FileInputStream(fileName);
        XSSFWorkbook workbook = new XSSFWorkbook(fis);

        XSSFSheet sheet = workbook.getSheet("Sheet1");
        Row rowField = sheet.getRow(rowNo);
        System.out.println(rowNo);

        Cell cellField = rowField.getCell(columnNo);
        System.out.println(columnNo);
        
        cellField.setCellValue(Value);

        FileOutputStream fos = new FileOutputStream(fileName);
        workbook.write(fos);
        workbook.close();
        fis.close();

    }


        
}
