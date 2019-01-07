package process;

import entity.Student;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.*;

public class XlsxBase {

    private Map<String, Student> xlsxData = new HashMap<>();

    public void readXlsx(String myDirectory) throws IOException {
        InputStream fileInputStream;
        XSSFWorkbook xssfWorkbook;
        XSSFSheet sheet;

// Put these loc in a try-catch block
        File dir = new File(myDirectory);
        File[] directoryListing = dir.listFiles();
        if (directoryListing != null) {
            for (File child : directoryListing) {

                fileInputStream = new FileInputStream(child);
                xssfWorkbook = new XSSFWorkbook(fileInputStream);
                sheet = xssfWorkbook.getSheetAt(0);
                iterateXlsRow(sheet);
            }
        }
    }

    private void iterateXlsRow(XSSFSheet sheet) {

        Iterator rowIterator;
        XSSFRow row;
        rowIterator = sheet.rowIterator();
        String time;
        String name;
        String classNumber;
        String reason;
        Random rand = new Random();
        DataFormatter formatter = new DataFormatter();

        // sheet.getRow(0);  why?
        while (rowIterator.hasNext()) {
            row = (XSSFRow) rowIterator.next();

            if (row.getCell(0) == null
                    || row.getCell(1) == null
                    || row.getCell(2) == null
                    || row.getCell(3) == null) {
                System.out.println("One of the cells are contains no data in row: " + row.getRowNum());
            }
            time = formatter.formatCellValue(row.getCell(0));
            name = row.getCell(1).getStringCellValue();
            classNumber = row.getCell(2).getStringCellValue();
            reason = row.getCell(3).getStringCellValue();

            String studentKey = row.getCell(1).getStringCellValue()
                    + "," + row.getCell(2).getStringCellValue() + "," + rand.nextInt(1000000) + 1;
            Student student = new Student(time, name, classNumber, reason);

            xlsxData.put(studentKey, student);
        }
    }

    public void writeXlsx() {
        Map<String, Student> map = new TreeMap<>(xlsxData);

    }

    /**
     * For testing
     * @return test result
     */
    public Map<String, Student> getXlsxData() {
        Map<String, Student> map = new TreeMap<>(xlsxData);
        return map;
    }

    public void stuff() {

        InputStream fileInputStream;
        HSSFWorkbook hssfWorkbook;
        HSSFSheet sheet = null;
        Iterator rowIterator, cellIterator;
        HSSFRow row = null;
        HSSFCell cell;
        rowIterator = sheet.rowIterator();
        cellIterator = row.cellIterator();

        while (cellIterator.hasNext()) {
            cell = (HSSFCell) cellIterator.next();
            if (cell.getCellTypeEnum() == CellType.STRING) {
                String someVariable = cell.getStringCellValue();
            } else if (cell.getCellTypeEnum() == CellType.NUMERIC) {
                // Handle numeric type
            } else {
                // Handle other types
            }
        }
        // Other code
    }
}
