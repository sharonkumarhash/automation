package excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Utils {
	
    public static void main(String[] args) throws Exception {
    	
        try {

            FileInputStream fileInputStream=new FileInputStream(new File("./Excel.xlsx"));
            XSSFWorkbook xssfWorkbook=new XSSFWorkbook(fileInputStream);
            XSSFSheet sheet=xssfWorkbook.getSheet("ObjectRepository");

            Iterator<Row> rowIterator=sheet.iterator();
            while (rowIterator.hasNext())
            {
                Row row = (Row)rowIterator.next();
                System.out.println("Row Number  "+row.getRowNum());
                Iterator<Cell> cellIterator=row.cellIterator();
                while (cellIterator.hasNext())
                {
                    Cell cell=(Cell)cellIterator.next();
                    System.out.print(cell.getStringCellValue() + ",");
                }
                System.out.println("");
                System.out.println("1");
            }

            fileInputStream.close();
            xssfWorkbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
	
		
}
