package eu.barjak.study_xlsx;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ToXLSX implements GlobalVariables {

    XSSFWorkbook wb = new XSSFWorkbook();
    int rowsize;

    public void write() throws FileNotFoundException, IOException {

        for (String sheetName : MAP.keySet()) {
            XSSFSheet sheet = wb.createSheet(sheetName);
            int r = 0;
            for (String key : MAP.get(sheetName).keySet()) {
                if (r == 0) {
                    rowsize = MAP.get(sheetName).get(key).size();
                }
                XSSFRow row = sheet.createRow(r++);
                for (int c = 0; c < rowsize; c++) {
                    row.createCell(c).setCellValue((String) MAP.get(sheetName).get(key).get(c));
                }
            }
        }
    }

    public void writeout() throws FileNotFoundException, IOException {
        String excelFileName = "out.xlsx";
        FileOutputStream fileOut = new FileOutputStream(excelFileName);
        wb.write(fileOut);
        fileOut.flush();
        fileOut.close();
    }
}
