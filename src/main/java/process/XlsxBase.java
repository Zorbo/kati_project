package process;

import entity.Student;

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
import java.util.Iterator;
import java.util.List;
import java.util.Vector;
import java.util.Comparator;


public class XlsxBase {

    private List<Student> xlsxDataList = new Vector<>();
    private String directory = "";
    private String year = "";

    /**
     * Read the files in the specific folder
     * @param year The year for the folder
     * @param month The month of the folder
     * @throws IOException If something goes wrong
     */
    public void readXlsx(String year, String month) throws IOException {
        XSSFWorkbook xssfWorkbook;
        XSSFSheet sheet;
        File dir = new File("C:\\Munka\\" + year + month);
        File[] directoryListing = dir.listFiles();
        if (directoryListing != null) {
            for (File child : directoryListing) {
                try(InputStream fileInputStream = new FileInputStream(child)) {
                    xssfWorkbook = new XSSFWorkbook(fileInputStream);
                    sheet = xssfWorkbook.getSheetAt(0);
                    iterateXlsRow(sheet);
                }
            }
        }
        this.directory = month;
        this.year = year;
    }

    /**
     * Iterate through the sheet for Student objects
     * @param sheet the sheet what we would like to read
     */
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
            Student student = new Student(time, name, classNumber, reason, classNumber + name);
            xlsxDataList.add(student);
        }
    }

    /**
     * Write out the xlsx file
     * @throws IOException If writing goes wrong
     */
    public void writeXlsx() throws IOException {
        String[] columns = {"Kártya lehúzása", "Tanuló neve", "Osztálya", "Késés oka"};
        xlsxDataList.sort(Comparator.comparing(Student::getKey));
        Workbook workbook = new XSSFWorkbook();
        CreationHelper creationHelper = workbook.getCreationHelper();
        Sheet sheet = workbook.createSheet("Munka1");

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
        for (Student student : xlsxDataList) {
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(student.getTime());
            row.createCell(1).setCellValue(student.getName());
            row.createCell(2).setCellValue(student.getClassNumber());
            row.createCell(3).setCellValue(student.getReason());
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
    public List<Student> getXlsxDataList() {
        xlsxDataList.sort(Comparator.comparing(Student::getKey));
        return xlsxDataList;
    }
}
