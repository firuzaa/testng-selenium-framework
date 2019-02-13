package excelAutomation;

import org.apache.poi.ss.usermodel.*;
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.util.HashMap;
import java.util.Map;

public class ExcelReadDemo {

    @Test
    public void readExcelFile() throws Exception {
        String path = "src/test/resources/country.xlsx";

        //Open file and convert to stream of data
        FileInputStream fileInputStream = new FileInputStream(path);

        //workbook > worksheet > row > cell
        //Open the workBook
        Workbook workbook = WorkbookFactory.create(fileInputStream);

        //Goto the first worksheet
        Sheet workSheet = workbook.getSheetAt(0);

        //Goto the first row
        Row row = workSheet.getRow(0);

        //Goto first cell
        Cell cell_1 = row.getCell(0);
        Cell cell_2 = row.getCell(1);
        //print cell values
        System.out.println(cell_1.toString());
        System.out.println(cell_2.toString());

        //read cell value using method chaining
        String country1 = workSheet.getRow(1).getCell(0).toString();
        String capital1 = workbook.getSheetAt(0).getRow(1).getCell(1).toString();

        System.out.println("Country 1: " + country1);
        System.out.println("Capital 1: " + capital1);

        int rowsCount = workSheet.getLastRowNum();
        System.out.println("Number of rows: " + rowsCount);
        System.out.println("********************************************");
        for (int i = 1; i <= rowsCount; i++) {
            System.out.println("Country #" + i + ": " + workSheet.getRow(i).getCell(0).toString() + " --> " + workSheet.getRow(i).getCell(1));
        }

        System.out.println("*******************************************");
        Map<String, String> countryMap = new HashMap<>();
        for (int x = 1; x <= rowsCount; x++) {
            countryMap.put(workSheet.getRow(x).getCell(0).toString(),workSheet.getRow(x).getCell(1).toString());
            System.out.println(countryMap);
        }

        System.out.println("*********************************************");
        Map<String, String> countriesMap = new HashMap<>();
        int countryCol = 0;
        int capitalCol = 1;

        for (int rowNum = 1; rowNum <= rowsCount;rowNum++){
            String country = workSheet.getRow(rowNum).getCell(countryCol).toString();
            String capital = workSheet.getRow(rowNum).getCell(capitalCol).toString();

            countriesMap.put(country,capital);
        }

        System.out.println(countriesMap);



        //close work book and steam
        workbook.close();
        fileInputStream.close();
    }
}
