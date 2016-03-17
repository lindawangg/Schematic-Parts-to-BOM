import java.io.*;
import java.util.*;

import org.apache.poi.hssf.util.HSSFColor;
//import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

public class CompareSheets {

	public static void main(String[] args) throws IOException {
		//Import the exported schematic spreadsheet
		Scanner input = new Scanner(System.in);
		System.out.println("Enter the Excel filename below:");
		String filename = input.nextLine();
		//String filename = "Book1.xlsx";
		File f = new File(filename);
		if(f.isFile() && f.exists())
		{
			System.out.println(filename + " file opened successfully.");
		}
		else
		{
			System.out.println("Error");
		}
		
		FileInputStream fis = null;
		try {
			fis = new FileInputStream(f);
		} catch (FileNotFoundException e1) {
			e1.printStackTrace();
		}
		XSSFWorkbook workbook = null;
		try {
			workbook = new XSSFWorkbook(fis);
		} catch (IOException e) {
			e.printStackTrace();
		}
	    XSSFSheet mSheet = workbook.getSheetAt(0);
	    XSSFSheet sSheet = workbook.getSheetAt(1);
	    
	    System.out.println("Enter the start row of PCBA sheet:");
	    int mstartRow = input.nextInt(); //5
	    input.nextLine();
	    System.out.println("Enter the value column:");
	    int mvalCol = input.nextInt(); //13
	    input.nextLine();
	    System.out.println("Enter the Ref Designator column:");
	    int mrefCol = input.nextInt(); //16
	    input.nextLine();
	    int sstartRow = 1;
	    int svalCol = 2;
	    int srefCol = 5;
	    input.close();
	    
	    //Open txt file writers 
	    PrintWriter toHighlight = null;
	    try {
			toHighlight = new PrintWriter("toHighlight.txt", "UTF-8");
		} catch (FileNotFoundException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		} catch (UnsupportedEncodingException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
	    
	    PrintWriter writer = null;
	    try {
			writer = new PrintWriter("result.txt", "UTF-8");
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (UnsupportedEncodingException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	    writer.println("Results");
	    
		//import manufacturer sheet into hashtable 
		System.out.println("Importing to hashtable");
		Hashtable<String, String> mtable = importSheetToHashtable(mSheet, mstartRow, mvalCol, mrefCol);
		//Hashtable<String, String> sTable = new Hashtable<String, String>();
		
		//writer.println(mtable);
		//System.out.println(mtable);
		
		//For each part reference in exported sheet, check if it is in table
		XSSFRow sRow = null;
		XSSFCell svalCell = null;
		XSSFCell srefCell;
		String sVal;
		String sRef;
		
		System.out.println("Comparing sheets");
		writer.println();
		String stringFormat = "%1$-10s %2$-15s %3$-20s";
		writer.println(String.format(stringFormat, "PART", "Action", "Reason"));
		writer.println("-----------------------------------------------");
		for (int r = sstartRow; r < sSheet.getPhysicalNumberOfRows(); r++) {
			sRow = sSheet.getRow(r);
			svalCell = sRow.getCell(svalCol);
			srefCell = sRow.getCell(srefCol);
			sVal = svalCell.getStringCellValue().toUpperCase();
			sRef = srefCell.getStringCellValue().toUpperCase();
			//System.out.println("Part Ref: " + sRef + " Value: " + sVal);
			
			if (sVal.contains("/NC") || sVal.contains("NC/") || sVal.equals("NC") || sRef.contains("TP") || sRef.contains("SH") || sRef.contains("PAD")) {
				if (mtable.containsKey(sRef)) {
					//The part is NC so should not be in mSheet
					toHighlight.println(sRef + " DELETE");
					writer.println(String.format(stringFormat, sRef, "Delete", "Part is NC in schematic"));
				}
				//The part is NC and in not in mSheet --> do nothing 
			}
			else {
				if (!mtable.containsKey(sRef)) {
					//The part should be in mSheet but is not 
					toHighlight.println(sRef + " ADD");
					//sTable.put(sRef, sVal);
					writer.println(String.format(stringFormat, sRef, "Add", "Part is not in sheet"));
				}
				else if (!mtable.get(sRef).contains("NA") && !ifEqual(mtable.get(sRef), sVal)) {
					//The part value in sSheet and mSheet do not equal 
					toHighlight.println(sRef + " CHECK");
					writer.println(String.format(stringFormat, sRef, "Check", "Values do not equal"));
					mtable.remove(sRef);
				}
				else {
					//The part in sSheet and mSheet are correct
					mtable.remove(sRef);
				}
			}	
		}
		
		//Check if hashtable is empty 
		writer.println();
		if (!mtable.isEmpty()) {
			writer.println("These parts were not found in schematic sheet.");
			writer.println(mtable);
		}
		
		//Close streams 
		writer.close();
		toHighlight.close();
		try {
			fis.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		//Call editSheets method
		System.out.println("Highlighting sSheet");
		
		BufferedReader reader = new BufferedReader(new FileReader("toHighlight.txt"));
		String line = "";
		line = reader.readLine();
		String[] parts = convertToParts(line);
		
		for (int r = sstartRow; r < sSheet.getPhysicalNumberOfRows(); r++) {
			sRow = sSheet.getRow(r);
			svalCell = sRow.getCell(svalCol);
			srefCell = sRow.getCell(srefCol);
			if (srefCell.getStringCellValue().equals(parts[0])) {
				if (parts[1].contains("DEL")){
					svalCell.setCellStyle(redstyle(workbook));
					srefCell.setCellStyle(redstyle(workbook));
					line = reader.readLine();
					if (line == null || line.length() < 3) break;
					parts = convertToParts(line);
				}
				else if (parts[1].contains("CHECK")) {
					svalCell.setCellStyle(yellowstyle(workbook));
					srefCell.setCellStyle(yellowstyle(workbook));
					line = reader.readLine();
					if (line == null || line.length() < 3) break;
					parts = convertToParts(line);
				}
				else if (parts[1].contains("ADD")) {
					svalCell.setCellStyle(bluestyle(workbook));
					srefCell.setCellStyle(bluestyle(workbook));
					line = reader.readLine();
					if (line == null || line.length() < 3) break;
					parts = convertToParts(line);
				}
			}
		}
		
		reader.close();
		FileOutputStream out = new FileOutputStream(new File("Book1.xlsx"));
		workbook.write(out);
		out.close();
		
		try {
			workbook.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		System.out.println("Finished");
	}
	
	public static XSSFCellStyle redstyle(XSSFWorkbook wb) {
		XSSFCellStyle redStyle = wb.createCellStyle();
		redStyle.setFillBackgroundColor(HSSFColor.RED.index);
		redStyle.setFillPattern(XSSFCellStyle.LEAST_DOTS);
		return redStyle;
	}
	
	public static XSSFCellStyle bluestyle(XSSFWorkbook wb) {
		XSSFCellStyle blueStyle = wb.createCellStyle();
		blueStyle.setFillBackgroundColor(HSSFColor.BLUE.index);
		blueStyle.setFillPattern(XSSFCellStyle.LEAST_DOTS);
		return blueStyle;
	}
	
	public static XSSFCellStyle yellowstyle(XSSFWorkbook wb) {
		XSSFCellStyle yellowStyle = wb.createCellStyle();
		yellowStyle.setFillBackgroundColor(HSSFColor.YELLOW.index);
		yellowStyle.setFillPattern(XSSFCellStyle.LEAST_DOTS);
		return yellowStyle;
	}
	
	public static String[] convertToParts(String str) {
		String[] parts = str.split(" ");
		/*
		for (int i = 0; i < parts.length; i++)
		{
			System.out.println(parts[i]);
		}
		*/
		return parts;
	}
	
	//Imports the Part Reference and Value of manufacturer sheet into Hashtable
	public static Hashtable<String, String> importSheetToHashtable(XSSFSheet s, int startRow, int valCol, int refCol) {
		Hashtable<String, String> t = new Hashtable<String, String>();
		//Iterate through rows on the excel sheet and put the part reference
		//and value into the hashtable 
		XSSFRow row;
		XSSFCell valCell;
		XSSFCell refCell;
		for (int r = startRow; r < s.getPhysicalNumberOfRows(); r++)
		{
			row = s.getRow(r);
			valCell = row.getCell(valCol);
			valCell.setCellType(XSSFCell.CELL_TYPE_STRING);
			refCell = row.getCell(refCol);
			refCell.setCellType(XSSFCell.CELL_TYPE_STRING);
			String ref = refCell.getStringCellValue().toUpperCase();
			String[] refValue = convertToParts(ref);
			//System.out.println("Ref Value " + ref);
			String valValue = valCell.getStringCellValue().toUpperCase();
			//System.out.println("Val Value " + valValue);
			for (int n = 0; n < refValue.length; n++){
				if (refValue[n] != "NA" && !refValue[n].isEmpty()) {
					if (valValue.contains("OHM")) {
						valValue = "NA";
					}
					t.put(refValue[n], valValue);
				}
			}
		}
	    return t;
	}
	
	public static boolean ifEqual(String mVal, String sVal) {
		if (sVal.length() == 0 || mVal.length() == 0) {
			if (sVal.equals(mVal))
				return true;
			else
				return false;
		}
		char[] char1 = mVal.toCharArray();
		char[] char2 = sVal.toCharArray();
		int match = 0;
		int j = 0;
		//System.out.println("mVal.length: " + mVal.length());
		for (int i = 0; i < char1.length; i++) {
			//System.out.println("Char1: " + char1[i] + " Char2: " + char2[j]);
			if (char1[i] == char2[j]) {
				match++;
				j++;
				if (j >= sVal.length())
					break;
			}
		}
		if ((match/(double)mVal.length())*100 > 0.5 )
			return true;
		return false;
	}

	public static void writeToBook(XSSFWorkbook workbook) throws IOException { 
	      //XSSFCellStyle style6 = workbook.createCellStyle();
	      //style6.setFillBackgroundColor(HSSFColor.RED.index );
	      //style6.setFillPattern(XSSFCellStyle.LEAST_DOTS);
	      //style6.setAlignment(XSSFCellStyle.ALIGN_FILL);
	      //sheet.setColumnWidth(1,8000);
	      //cell.setCellStyle(style6);
	      
	      FileOutputStream out = new FileOutputStream(new File("Book1.xlsx"));
	      
	      workbook.write(out);
	      out.close();
	      //workbook.close();
	      System.out.println("Book1.xlsx written successfully");
	}
	
	public static void editSheets(String filename, XSSFWorkbook wb, int mStartRow, int mvalCol, int mrefCol, int sStartRow, int svalCol, int srefCol) throws IOException {
		//XSSFSheet mSheet = wb.getSheetAt(0);
		XSSFSheet sSheet = wb.getSheetAt(1);
		
		BufferedReader reader = new BufferedReader(new FileReader("toHighlight.txt"));
		
		String line = "";
		line = reader.readLine();
		String[] parts = convertToParts(line);
		XSSFRow sRow = null;
		XSSFCell svalCell =null;
		XSSFCell srefCell = null;
		
		XSSFCellStyle redStyle = wb.createCellStyle();
		redStyle.setFillBackgroundColor(HSSFColor.RED.index);
		redStyle.setFillPattern(XSSFCellStyle.LEAST_DOTS);
		
		XSSFCellStyle yellowStyle = wb.createCellStyle();
		yellowStyle.setFillBackgroundColor(HSSFColor.YELLOW.index);
		yellowStyle.setFillPattern(XSSFCellStyle.LEAST_DOTS);
		
		XSSFCellStyle blueStyle = wb.createCellStyle();
		blueStyle.setFillBackgroundColor(HSSFColor.BLUE.index);
		blueStyle.setFillPattern(XSSFCellStyle.LEAST_DOTS);
		
		FileOutputStream out = new FileOutputStream(new File("Book1.xlsx"));
	 
		for (int r = sStartRow; r < sSheet.getPhysicalNumberOfRows(); r++) {
			sRow = sSheet.getRow(r);
			svalCell = sRow.getCell(svalCol);
			srefCell = sRow.getCell(srefCol);
			if (srefCell.getStringCellValue().equals(parts[0])) {
				if (parts[1].contains("DEL")){
					svalCell.setCellStyle(redStyle);
					srefCell.setCellStyle(redStyle);
					wb.write(out);
					line = reader.readLine();
					parts = convertToParts(line);
				}
				else if (parts[1].contains("CHECK")) {
					svalCell.setCellStyle(yellowStyle);
					srefCell.setCellStyle(redStyle);
					line = reader.readLine();
					parts = convertToParts(line);
				}
				else if (parts[1].contains("ADD")) {
					svalCell.setCellStyle(blueStyle);
					srefCell.setCellStyle(redStyle);
					line = reader.readLine();
					parts = convertToParts(line);
				}
			}
		}
		reader.close();
		out.close();
	}
}
