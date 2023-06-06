package zadatak1;


import com.github.javafaker.Faker;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;


public class Zadatak1 {
    public static void main(String[] args) {

        String relativePath = "NameData.xlsx";

        try {
            readAndWriteData(relativePath);
        } catch (FileNotFoundException exception) {
            System.out.println("Invalid path");
        } catch (NullPointerException nullPointerException) {
            System.out.println("Data does not exist");
        } catch (IOException ioException) {
            System.out.println("Invalid Excel file");
        }
    }

    public static void readAndWriteData(String relativePath) throws IOException {

        FileInputStream fileInputStream = new FileInputStream(relativePath);
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet sheet = workbook.getSheet("First sheet");

        XSSFSheet sheet2 = workbook.createSheet("New sheet");

        XSSFRow row = sheet.getRow(0);
        XSSFCell cell = row.getCell(0);

        int rowCount = sheet.getPhysicalNumberOfRows();
        int cellCount = row.getPhysicalNumberOfCells();

        ArrayList<String> names = new ArrayList<>();

        for (int i = 0; i < rowCount; i++) {

            row = sheet.getRow(i);

            for (int j = 0; j < cellCount; j++) {
                
                cell = row.getCell(j);
                String name = cell.getStringCellValue();
                names.add(name);
            }
        }
        System.out.print(names); // first 5 names from first Excel sheet
        System.out.println();

        Faker faker = new Faker();

        for (int i = 0; i < 5; i++) {
            names.add(faker.name().firstName());
        }

        System.out.println(names);  //new 5 names generated using Faker

        for (int i = 0; i < names.size(); i++) {  //writing all 10 names in new sheet

            XSSFRow row2 = sheet2.createRow(i);

            for (int j = 0; j < cellCount; j++) {

                XSSFCell cell2 = row2.createCell(j);
                cell2.setCellValue(names.get(i));
            }
        }
        FileOutputStream fileOutputStream = new FileOutputStream(relativePath);
        workbook.write(fileOutputStream);
        fileOutputStream.close();
    }
}
