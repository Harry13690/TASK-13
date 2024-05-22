import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class Question_3 {

    public static void main(String[] args) throws IOException, InvalidFormatException {

        File f = new File("src/test/resources/ToWriteExcelData.xlsx");

        FileOutputStream fo = new FileOutputStream(f);
        
        XSSFWorkbook wb = new XSSFWorkbook();

        Sheet sh = wb.createSheet("Sheet1");

        Row columnhead = sh.createRow(0);
        columnhead.createCell(0).setCellValue("Name");
        columnhead.createCell(1).setCellValue("Age");
        columnhead.createCell(2).setCellValue("Email");

        Row r1 = sh.createRow(1);
        r1.createCell(0).setCellValue("John Doe");
        r1.createCell(1).setCellValue(30);
        r1.createCell(2).setCellValue("john@test.com");

        Row r2 = sh.createRow(2);
        r2.createCell(0).setCellValue("Jane Doe");
        r2.createCell(1).setCellValue(28);
        r2.createCell(2).setCellValue("john@test.com");

        Row r3 = sh.createRow(3);
        r3.createCell(0).setCellValue("Bob Smith");
        r3.createCell(1).setCellValue(35);
        r3.createCell(2).setCellValue("jacky@example.com");

        Row r4 = sh.createRow(4);
        r4.createCell(0).setCellValue("Swapnil");
        r4.createCell(1).setCellValue(37);
        r4.createCell(2).setCellValue("swapnil@example.com");
        
        wb.write(fo);
        
        wb.close();

        System.out.println("The Following Data are Inserted");
    }
}
