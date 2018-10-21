package factory;

import lombok.Data;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;

@Data
public abstract class XlsxBase {


    public void readXlsx(String path, String fileName) {

        try {
            POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(path + fileName));
            HSSFWorkbook wb = new HSSFWorkbook(fs);
            // Always read sheet position 0
            HSSFSheet sheet = wb.getSheetAt(0);
            HSSFRow row;
            HSSFCell cell;

            int rows; //no of rows
            rows = sheet.getPhysicalNumberOfRows();

            int cols = 0; //no of columns
            int tmp = 0;

            // This trick ensures that we get the data properly even if it doesn't start from first few rows
            for (int i = 0; i < 10 || i < rows; i++) {
                row = sheet.getRow(i);
                if (row != null) {
                    tmp = sheet.getRow(i).getPhysicalNumberOfCells();
                    if (tmp > cols) {
                        cols = tmp;
                    }
                }
            }

            for (int r = 0; r < rows; r++) {
                row = sheet.getRow(r);
                if (row != null) {
                    for (int c = 0; c < cols; c++) {
                        cell = row.getCell((short) c);
                        if (cell != null) {
                            // Your code here
                            cell.getStringCellValue(); //- for numeric cells we throw an exception.

                        }

                    }
                }
            }

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void readXlsxRowIterator(String path, String fileName) throws IOException {

        InputStream fileInputStream = null;
        HSSFWorkbook hssfWorkbook;
        HSSFSheet sheet;
        HSSFRow row;
        HSSFCell cell;
        Iterator rowIterator, cellIterator;

// Put these loc in a try-catch block
        fileInputStream = new FileInputStream(path + fileName);
        hssfWorkbook = new HSSFWorkbook(fileInputStream);
        sheet = hssfWorkbook.getSheetAt(0);
        rowIterator = sheet.rowIterator();

        // Get the rows from 0-3, BEÉRKEZÉS IDEJE, NEV, OSZTALY, KESES OKA

        sheet.getRow(0);

        while (rowIterator.hasNext()) {
            row = (HSSFRow) rowIterator.next();
            cellIterator = row.cellIterator();
            while (cellIterator.hasNext()) {
                cell = (HSSFCell) cellIterator.next();

                if (cell.getCellType() == HSSFCell.CELL_TYPE_STRING) {
                    String someVariable = cell.getStringCellValue();
                } else if (cell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC) {
                    // Handle numeric type
                } else {
                    // Handle other types
                }
            }
            // Other code
        }
    }

    // Implement XLSX writing
}
