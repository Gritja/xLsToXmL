package xLsToXmL;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileWriter;
import java.io.PrintWriter;
import java.text.*;

public class FileReader {
	// this is stupid, working on a simple syntax change to make this work better
	static DecimalFormat df = new DecimalFormat("#####0");

	public static void converterMethod(String excelFilePath, String filePath, String fileName) {
		FileWriter fostream;
		PrintWriter prStream;
		//hard-coded file names are the best! 
		//String fileName = "RenameMe";

		File excelFile = new File(excelFilePath);
		if (excelFile.exists()) {
			Workbook workbook = null;
			FileInputStream inputStream = null;

			try {
				inputStream = new FileInputStream(excelFile);
				workbook = new XSSFWorkbook(inputStream);
				Sheet sheet = workbook.getSheetAt(0);
				//handle creates and writes to a new file.
				fostream = new FileWriter(filePath + "\\" + fileName + ".xml");
				prStream = new PrintWriter(new BufferedWriter(fostream));

				prStream.println("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
				prStream.println("<WorkGroups>");
				//gotta have a group number! 
				int groupNumber = 0;
				boolean firstRow = true;
				int lastColumn = 0;
				int firstRowNumber = 0;
				//row iterator, goes through first row and ignores cells that contains nothing
				for (Row row : sheet) {
					if (firstRow) {
						lastColumn = row.getLastCellNum();
						firstRow = false;
						continue;
					}
					prStream.println("\t<Group no= \"" + groupNumber + "\">");
					//manually iterating through each column to check if there's supposed to be a value in this column
					   for (int cn = 0; cn < lastColumn; cn++) {												
						      Cell c = row.getCell(cn, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
						      Cell fc = sheet.getRow(firstRowNumber).getCell(cn, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
						      if (c == null) {
						         // The spreadsheet is empty in this cell
						    	  if (fc != null) {
						    		  prStream.println(formatElement("\t\t", formatCell(fc),
												formatCell(c)));
						    	  }
						    	  
						      } else {
					         // Do something useful with the cell's contents
						    	  prStream.println(formatElement("\t\t", formatCell(fc),
											formatCell(c)));
						      }

						   }
						groupNumber++;
						prStream.println("\t</Group>");
				}
				


				prStream.println("</WorkGroups>");
				// closes the writer
				prStream.close();
			} catch (IOException e) {
				// handle the exception
			} finally {
				IOUtils.closeQuietly(workbook);
				IOUtils.closeQuietly(inputStream);

			}

		}
	}
//formats the cell depending on it's contents.
	public static String formatCell(Cell cell) {
		if (cell == null) {
			return "";
		}
		switch (cell.getCellType()) {
		case BLANK:
			return "";
		case BOOLEAN:
			return Boolean.toString(cell.getBooleanCellValue());
		case ERROR:
			return "*error*";
		case NUMERIC:
			return FileReader.df.format(cell.getNumericCellValue());
		case STRING:
			//not too adept at xml format, so this uses regex to replace all symbols incompatible with xml tags
			return cell.getStringCellValue().replaceAll("[ #()?]", "");
		default:
			return "<unknown value>";
		}

	}

//a string building method I "borrowed", because it was nifty.
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