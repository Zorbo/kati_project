package run;

import entity.StudentInfo;
import process.XlsxBase;

import java.io.IOException;
import java.util.Map;

public class StartProcess {

    public static void main(String[] args) throws IOException {

        XlsxBase xlsxBase = new XlsxBase();

        xlsxBase.readXlsx("C:\\Users\\tamas.a.kiss\\Desktop\\Kati_Project\\xlsx_data\\October");
        xlsxBase.writeXlsx();

//        for (Map.Entry<String, StudentInfo> entry : xlsxBase.getXlsxData().entrySet()) {
//            System.out.println("Key: " + entry.getKey() + " Value: " + entry.getValue().getReason() + " Time: " + entry.getValue().getTime() + "\n");
//        }

    }
}
