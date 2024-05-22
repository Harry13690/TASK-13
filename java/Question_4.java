import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class Question_4 {

    public static void main(String[] args) throws IOException {

        File f = new File("src/test/resources/ToWriteExcelData.xlsx");

        FileOutputStream fo = new FileOutputStream(f);

        XSSFWorkbook wb = new XSSFWorkbook();

        Sheet sh = wb.createSheet("Sheet1");

        Row columnhead = sh.createRow(0);
        columnhead.createCell(0).setCellValue("Test Case");
        columnhead.createCell(1).setCellValue("Pass/Fail");
        columnhead.createCell(2).setCellValue("Priority");

        Row r1 = sh.createRow(1);
        r1.createCell(0).setCellValue("TC_1");
        r1.createCell(1).setCellValue("Pass");
        r1.createCell(2).setCellValue("High");

        Row r2 = sh.createRow(2);
        r2.createCell(0).setCellValue("TC_2");
        r2.createCell(1).setCellValue("Fail");
        r2.createCell(2).setCellValue("High");

        Row r3 = sh.createRow(3);
        r3.createCell(0).setCellValue("Tc_3");
        r3.createCell(1).setCellValue("Pass");
        r3.createCell(2).setCellValue("Low");

        wb.write(fo);

        wb.close();
        System.out.println("Values are Inserted");
    }
}
