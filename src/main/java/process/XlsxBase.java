package process;

import entity.StudentInfo;

import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.FileOutputStream;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.TreeMap;
import java.util.Random;


public class XlsxBase {

    private Map<String, StudentInfo> xlsxData = new HashMap<>();
    private String directory = "";
    private String year = "";

    public void readXlsx(String year, String month) throws IOException {
        InputStream fileInputStream;
        XSSFWorkbook xssfWorkbook;
        XSSFSheet sheet;

// Put these loc in a try-catch block
        File dir = new File("C:\\Munka\\" + year + month);
        File[] directoryListing = dir.listFiles();
        if (directoryListing != null) {
            for (File child : directoryListing) {

                fileInputStream = new FileInputStream(child);
                xssfWorkbook = new XSSFWorkbook(fileInputStream);
                sheet = xssfWorkbook.getSheetAt(0);
                iterateXlsRow(sheet);
            }
        }
        this.directory = month;
        this.year = year;
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
            String studentKey = classNumber + "," + name + "," + time;
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
        String studentName = "";
        String studentClassNumber = "";

        // Create header
        Font headerFont = workbook.createFont();
        headerFont.setFontHeightInPoints((short) 14);
        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);
        headerCellStyle.setBorderBottom(BorderStyle.MEDIUM);
        headerCellStyle.setBorderTop(BorderStyle.MEDIUM);

        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < columns.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(columns[i]);
            cell.setCellStyle(headerCellStyle);
        }

        // Create Other rows and cells with student data
        int rowNum = 1;
        Random rand = new Random();
        for (StudentInfo studentInfo : map.values()) {
            Row row = sheet.createRow(rowNum++);

            if (studentInfo.getName().equals(studentName)
                    && studentInfo.getClassNumber().equals(studentClassNumber)) {
                // Set student row color
                Font rowFont = workbook.createFont();
                rowFont.setColor((short) (rand.nextInt(10) + 1));
                CellStyle rowCellStyle = workbook.createCellStyle();
                rowCellStyle.setFont(rowFont);
            }
            row.createCell(0).setCellValue(studentInfo.getTime());
            row.createCell(1).setCellValue(studentInfo.getName());
            row.createCell(2).setCellValue(studentInfo.getClassNumber());
            row.createCell(3).setCellValue(studentInfo.getReason());

            studentName = studentInfo.getName();
            studentClassNumber = studentInfo.getClassNumber();
        }
        // Resize all columns to fit the content size
        for (int i = 0; i < columns.length; i++) {
            sheet.autoSizeColumn(i);
        }

        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream("C:\\Munka\\" + year + "_" + directory + "_ho_vegi_osszesites.xlsx");
        workbook.write(fileOut);
        fileOut.close();
        // Closing the workbook
        workbook.close();
    }

    /**
     * For testing
     *
     * @return test result
     */
    public Map<String, StudentInfo> getXlsxData() {
        return new TreeMap<>(xlsxData);
    }
}
