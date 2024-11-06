package hcyclewala;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class dataDriven 
{
    public ArrayList<String> getData(String testCase) throws IOException{
        FileInputStream fis = new FileInputStream(System.getProperty("user.dir") + "\\exceldatadriven\\ExcelData.xlsx");
        //workbook
        //excel-data-driven
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        ArrayList<String> a = new ArrayList<String>();
        int sheets = workbook.getNumberOfSheets();
        for(int i=0; i<sheets; i++)
        {
            if (workbook.getSheetName(i).equalsIgnoreCase("data"))
            {
                XSSFSheet sheet = workbook.getSheetAt(i);

                Iterator<Row> rows = sheet.iterator();
                Row firstrow = rows.next();
                
                Iterator<Cell> ce = firstrow.cellIterator();

                int k=0;
                int column =0;
                while(ce.hasNext())
                {
                    Cell value = ce.next();
                    if(value.getStringCellValue().equalsIgnoreCase("testcases"))
                    {
                        column = k;
                    }
                    k++;
                }
                System.out.println(column);

                while(rows.hasNext())
                {
                    Row r = rows.next();
                    if(r.getCell(column).getStringCellValue().equalsIgnoreCase(testCase))
                    {
                        Iterator<Cell> c = r.cellIterator();
                        while (c.hasNext()) {
                            Cell v = c.next();
                            if(v.getCellType()==CellType.STRING)
                            {
                            a.add(v.getStringCellValue());
                            }
                            else{
                                a.add(NumberToTextConverter.toText(v.getNumericCellValue()));
                            }
                        }
                    }
                }
            }
        }
        System.out.println(a);
        return a;
    }
}
