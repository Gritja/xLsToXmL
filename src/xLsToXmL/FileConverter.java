package xLsToXmL;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintWriter;
import java.text.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class FileConverter {
	static DecimalFormat df = new DecimalFormat("#####0");
public static void converterMethod() {
	FileWriter fostream;
    PrintWriter out = null;
    String strOutputPath = "C:\\Users\\Jonatan\\Dropbox\\mitt\\java19\\javaprog\\";
    String strFilePrefix = "testing";

    try {
        InputStream inputStream = new FileInputStream(new File("SheettobecomeXML.xlsx"));
        Workbook wb = WorkbookFactory.create(inputStream);
        Sheet sheet = wb.getSheet("Bin-code");

        fostream = new FileWriter(strOutputPath + "\\" + strFilePrefix+ ".xml");
        out = new PrintWriter(new BufferedWriter(fostream));

        out.println("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
        out.println("<Bin-code>");

        boolean firstRow = true;
        for (Row row : sheet) {
            if (firstRow == true) {
                firstRow = false;
                continue;
            }
            out.println("\t<DCT>");
            out.println(formatElement("\t\t", "ID", formatCell(row.getCell(0))));
            out.println(formatElement("\t\t", "Variable", formatCell(row.getCell(1))));
            out.println(formatElement("\t\t", "Desc", formatCell(row.getCell(2))));
            out.println(formatElement("\t\t", "Notes", formatCell(row.getCell(3))));
            out.println("\t</DCT>");
        }
        out.write("</Bin-code>");
        out.flush();
        out.close();
    } catch (Exception e) {
        e.printStackTrace();
    }
    out.close();
}

public static String formatCell(Cell cell)
{
    if (cell == null) {
        return "";
    }
    switch(cell.getCellType()) {
        case BLANK:
            return "";
        case BOOLEAN:
            return Boolean.toString(cell.getBooleanCellValue());
        case ERROR:
            return "*error*";
        case NUMERIC:
            return FileConverter.df.format(cell.getNumericCellValue());
        case STRING:
            return cell.getStringCellValue();
        default:
            return "<unknown value>";
    }

}

public static String formatElement(String prefix, String tag, String value) {
    StringBuilder sb = new StringBuilder(prefix);
    sb.append("<");
    sb.append(tag);
    if (value != null && value.length() > 0) {
        sb.append(">");
        sb.append(value);
        sb.append("</");
        sb.append(tag);
        sb.append(">");
    } else {
        sb.append("/>");
    }
    return sb.toString();
}

}