import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class Question_1 {

    private static void createNewWorkBook() throws IOException, InvalidFormatException {
        File f = new File("src/test/resources/ToWriteExcelData.xlsx");

        FileOutputStream fo = new FileOutputStream(f);

        XSSFWorkbook wb = new XSSFWorkbook();

        wb.write(fo);

        wb.close();

        System.out.println("New WorkBook Is Created Successfully");
    }

    public static void main(String[] args) throws IOException, InvalidFormatException {

        createNewWorkBook();
    }
}