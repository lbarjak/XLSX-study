package eu.barjak.study_xlsx;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.LinkedHashMap;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class FromXLSX implements GlobalVariables {

    String sheetName;

    public void read(String xlsxName) throws FileNotFoundException, IOException, OpenXML4JException {

        OPCPackage fis = OPCPackage.open(new FileInputStream(xlsxName));
        XSSFReader r = new XSSFReader(fis);
        XSSFWorkbook myWorkBook = new XSSFWorkbook(fis);
        Iterator<InputStream> sheets = r.getSheetsData();

        XSSFReader.SheetIterator sheetiterator = (XSSFReader.SheetIterator) sheets;

        while (sheetiterator.hasNext()) {
            sheetiterator.next();
            sheetName = sheetiterator.getSheetName();
            MAP.put(sheetName, new LinkedHashMap<>());
            XSSFSheet actualSheet = myWorkBook.getSheet(sheetName);

            int numberOfColumns = actualSheet.getRow(0).getPhysicalNumberOfCells();
            for (Row rowOfSheet : actualSheet) {
                ArrayList<String> rowOfArlistaArrayList = new ArrayList<>();
                for (int c = 0; c < numberOfColumns; c++) {
                    rowOfArlistaArrayList.add(null);
                    Cell cell = rowOfSheet.getCell(c);
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
                MAP.get(sheetName).put(key, new ArrayList<>(rowOfArlistaArrayList));
            }
        }
        fis.close();
    }
}
