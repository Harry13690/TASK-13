import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class Question_2 {

    public static void createNewSheet() throws IOException, InvalidFormatException {
        File f = new File("src/test/resources/ToWriteExcelData.xlsx");

        FileOutputStream fo = new FileOutputStream(f);

        XSSFWorkbook wb = new XSSFWorkbook();

        Sheet sh = wb.createSheet("Sheet1");

        wb.write(fo);

        wb.close();

        System.out.println("New Sheet Is Created Successfully");
    }

    public static void main(String[] args) throws IOException, InvalidFormatException {

        createNewSheet();
    }
}
