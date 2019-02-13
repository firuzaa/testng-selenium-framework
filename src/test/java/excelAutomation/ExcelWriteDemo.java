package excelAutomation;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.io.FileOutputStream;

public class ExcelWriteDemo {

    @Test
    public void writeExcel() throws Exception {

        // "." means go to this root folder in this project
        String filePath = "./src/test/resources/country.xlsx";
        FileInputStream fileInputStream = new FileInputStream(filePath);
        Workbook workbook = WorkbookFactory.create(fileInputStream);
        Sheet sheet = workbook.getSheetAt(0);

        //Write column name

        Cell cell = sheet.getRow(0).getCell(2);
        if (cell == null) {
            cell = sheet.getRow(0).createCell(2);
        }
        cell.setCellValue("Continent");

        Cell cell1 = sheet.getRow(1).createCell(2);
        if (cell1 == null) {
            cell1 = sheet.getRow(1).createCell(2);
        }
        cell1.setCellValue("North America");

        //Save changes
        //Open the file to WRITE into it
        FileOutputStream fileOutputStream = new FileOutputStream(filePath);

        //Write and save the changes
        workbook.write(fileOutputStream);

        fileOutputStream.close();
        workbook.close();
        fileInputStream.close();


    }

}
