import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;

public class Question_5 {

    private static void toPrintData() throws IOException {

        FileInputStream fi = new FileInputStream("src/test/resources/ToReadExcelData.xlsx");

        XSSFWorkbook wb = new XSSFWorkbook(fi);

        Sheet sh = wb.getSheet("Sheet1");

        for (int i = 0; i < sh.getPhysicalNumberOfRows(); i++) {
            Row r = sh.getRow(i);
            for (int j = 0; j < r.getPhysicalNumberOfCells(); j++) {
                Cell cell = r.getCell(j);
                System.out.println(cell);
            }
        }
    }

    public static void main(String[] args) throws IOException {

        toPrintData();
    }
}
