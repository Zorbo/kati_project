package process;

import entity.StudentInfo;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class XlsxBase {

    private Map<String, StudentInfo> xlsxData = new HashMap<>();

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
            String studentKey = name + "," + classNumber + "," + time;

            StudentInfo studentInfo = new StudentInfo(time, name, classNumber, reason);

            xlsxData.put(studentKey, studentInfo);
        }
    }

    public void writeXlsx() throws IOException {

        String[] columns = {"Kártya lehúzása", "Tanuló neve", "Osztálya", "Késés oka"};
        Map<String, StudentInfo> map = new TreeMap<>(xlsxData);
        Workbook workbook = new XSSFWorkbook();
        CreationHelper creationHelper = workbook.getCreationHelper();
        Sheet sheet = workbook.createSheet("Munka1");

        // Create header
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 14);
        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);

        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < columns.length; i++){
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(columns[i]);
            cell.setCellStyle(headerCellStyle);
        }

        // Create Other rows and cells with student data
        int rowNum = 1;
        for (StudentInfo studentInfo : map.values()){
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(studentInfo.getTime());
            row.createCell(1).setCellValue(studentInfo.getName());
            row.createCell(2).setCellValue(studentInfo.getClassNumber());
            row.createCell(3).setCellValue(studentInfo.getReason());
        }

        // Resize all columns to fit the content size
        for(int i = 0; i < columns.length; i++) {
            sheet.autoSizeColumn(i);
        }

        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream("C:\\Users\\tamas.a.kiss\\Desktop\\Kati_Project\\xlsx_data\\October\\poi-generated-file.xlsx");
        workbook.write(fileOut);
        fileOut.close();

        // Closing the workbook
        workbook.close();

    }

    /**
     * For testing
     * @return test result
     */
    public Map<String, StudentInfo> getXlsxData() {
        return new TreeMap<>(xlsxData);
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
