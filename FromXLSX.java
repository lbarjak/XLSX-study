package eu.barjak.study_xlsx;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class FromXLSX implements GlobalVariables {

    XSSFWorkbook myWorkBook;
    String sheetName;

    public void read(String xlsxName) throws FileNotFoundException, IOException {

        File myFile = new File(xlsxName);
        FileInputStream fis = new FileInputStream(myFile);
        myWorkBook = new XSSFWorkbook(fis);
        int numOfSheets = myWorkBook.getNumberOfSheets();
        
        for (int sheetNumber = 0; sheetNumber < numOfSheets; sheetNumber++) {
            
            String sheetCount = myWorkBook.getSheetName(sheetNumber);
            SHEET_NAMES.add(sheetCount);
            MAP.put(sheetCount, new LinkedHashMap<>());
            XSSFSheet mySheet = myWorkBook.getSheetAt(sheetNumber);
            int numberOfColumns = mySheet.getRow(0).getPhysicalNumberOfCells();
            for (Row rowOfWorkbook : mySheet) {

                ArrayList<String> rowOfArlistaArrayList = new ArrayList<>();

                for (int c = 0; c < numberOfColumns; c++) {
                    rowOfArlistaArrayList.add(null);
                    Cell cell = rowOfWorkbook.getCell(c);
                    if (!(cell == null)) {
                        switch (cell.getCellType()) {
                            case Cell.CELL_TYPE_STRING:
                                rowOfArlistaArrayList.set(c, cell.getStringCellValue().trim());
                                break;
                            case Cell.CELL_TYPE_NUMERIC:
                                rowOfArlistaArrayList.set(c, String.valueOf(cell.getNumericCellValue()));
                                break;
                            case Cell.CELL_TYPE_BOOLEAN:
                                rowOfArlistaArrayList.set(c, String.valueOf(cell.getBooleanCellValue()));
                                break;
                            default:
                        }
                    }
                }
                String key = rowOfArlistaArrayList.get(0);
                MAP.get(sheetCount).put(key, new ArrayList<>(rowOfArlistaArrayList));
            }
        }
    }
}
