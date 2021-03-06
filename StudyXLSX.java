package eu.barjak.study_xlsx;

import java.io.FileNotFoundException;
import java.io.IOException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;

//-Xmx1792m - -Xmx3072m Heap size
public class StudyXLSX implements GlobalVariables {

    public static String xlsxName = "sample.xlsx";

    public static void main(String[] args) throws FileNotFoundException, IOException, OpenXML4JException {

        FromXLSX fromxlsx = new FromXLSX();
        ToXLSX toxlsx = new ToXLSX();

        fromxlsx.read(xlsxName);
        
        toxlsx.write();
        toxlsx.writeout();
    }
}
