package com.Feasible;

import java.io.File; 
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class SynchDataPools2 {
	static XSSFWorkbook Masterworkbook;

//	public static void main(String[] args) throws IOException {
//		String RootPath = "D:\\testFiles\\";
//		copyMaster(RootPath);
//	}protected

	protected void copyMaster(String RootPath) throws IOException {

		XSSFSheet mappingsheet;
		// String FilePath = System.getProperty("user.dir")
		// + "\\src\\test\\resources\\MasterFile.xls";

		String FilePath = RootPath + "MasterFile.xlsx";

		FileInputStream inputFile = new FileInputStream(new File(FilePath));
		Masterworkbook = new XSSFWorkbook(inputFile);
		// mappingsheet creation
		mappingsheet = Masterworkbook.getSheet("MappingFile");

		if (mappingsheet.equals(null)) {
			System.out.println(" \n no such sheet is there");
		} else {
			int noOfColumns = mappingsheet.getRow(3).getLastCellNum();
			// capture child Files

			for (int i = 1; i < noOfColumns; i++) {

				String str = String.valueOf(mappingsheet.getRow(3).getCell(i));
				String[] splittedStr = str.split(":");
				FilePath = null;

				System.out.println(" \n parent workbook :	" + splittedStr[0]);
				FilePath = RootPath + splittedStr[0];
				//FilePath = splittedStr[0];
				FileInputStream childInputFile = new FileInputStream(new File(
						FilePath));
				// child workbook creation
				XSSFWorkbook childworkbook;
				XSSFSheet childsheet;

				childworkbook = new XSSFWorkbook(childInputFile);
				String childsheetName = splittedStr[1];
				// Child sheet creation
				childsheet = childworkbook.getSheet(childsheetName);
				int Map_lastrow = mappingsheet.getPhysicalNumberOfRows() + 3;
				System.out.println(" \n Physical Number Of Rows :"
						+ Map_lastrow);
				for (int row = 4; row < Map_lastrow; row++) {

					if (mappingsheet.getRow(row).getCell(i).getCellType() != XSSFCell.CELL_TYPE_BLANK) {
						System.out
								.println(" \nParentsheet -  value in column :"
										+ i + "\t and row :" + row
										+ "\t is   : "
										+ mappingsheet.getRow(row).getCell(i));
						str = null;
						str = String.valueOf(mappingsheet.getRow(row)
								.getCell(0));
						splittedStr = null;
						splittedStr = str.split(":");
						// sheet name and column number of the Master
						String parentsheetName = splittedStr[0];
						XSSFSheet parentsheet = Masterworkbook
								.getSheet(parentsheetName);
						int parentColumnNumber = Integer
								.parseInt(splittedStr[1]);

						int childColumnNumber = (int) mappingsheet.getRow(row)
								.getCell(i).getNumericCellValue();
						copycolumn(parentsheet, parentColumnNumber, childsheet,
								childColumnNumber);

					}

				}
				childInputFile.close();
				FileOutputStream outFile;
				try {
					outFile = new FileOutputStream(new File(FilePath));
					childworkbook.write(outFile);
					outFile.close();
					childworkbook.close();
				} catch (FileNotFoundException e) {
					e.printStackTrace();
				}
			}

		}
		inputFile.close();
		Masterworkbook.close();
	}

	private static void copycolumn(XSSFSheet parentsheet,
			int parentColumnNumber, XSSFSheet childsheet, int childColumnNumber) {
		// XSSFRow row=null;
		System.out
				.println("\n-------------------copycolumn()------------------------\n");
		System.out.println("parentsheet :" + parentsheet.getSheetName());
		System.out.println(" parentColumnNumber :" + parentColumnNumber);
		System.out.println("childsheet :" + childsheet.getSheetName());
		System.out.println(" childColumnNumber :" + childColumnNumber);

		int PSheetLastRowNumber = parentsheet.getPhysicalNumberOfRows();

		System.out.println("\n copying the columns...");
		for (int rownumber = 0; rownumber < PSheetLastRowNumber; rownumber++) {

			if (parentsheet.getRow(rownumber).getCell(parentColumnNumber) == null)
				continue;
			XSSFCell parentCell = parentsheet.getRow(rownumber).getCell(
					parentColumnNumber);
			XSSFRow childRow;
			XSSFCell childCell;

			if (childsheet.getRow(rownumber) != null)
				childRow = childsheet.getRow(rownumber);
			else
				childRow = childsheet.createRow(rownumber);

			if (childRow.getCell(childColumnNumber) != null)
				childCell = childRow.getCell(childColumnNumber);
			else
				childCell = childRow.createCell(childColumnNumber);

			System.out.println("column number : " + childColumnNumber
					+ " of Childsheet : " + childsheet.getSheetName());

			if (parentCell.getCellType() == XSSFCell.CELL_TYPE_STRING) {
				System.out.println("STRING Column");
				childCell.setCellType(XSSFCell.CELL_TYPE_STRING);
				childCell.setCellValue(parentCell.getStringCellValue());
			} else if (parentCell.getCellType() == XSSFCell.CELL_TYPE_NUMERIC) {
				System.out.println("NUMERIC Column");
				childCell.setCellType(XSSFCell.CELL_TYPE_NUMERIC);
				childCell.setCellValue(parentCell.getNumericCellValue());
			} else if (parentCell.getCellType() == XSSFCell.CELL_TYPE_BOOLEAN) {
				System.out.println("BOOLEAN Column");
				childCell.setCellType(XSSFCell.CELL_TYPE_BOOLEAN);
				childCell.setCellValue(parentCell.getBooleanCellValue());
			} else if (parentCell.getCellType() == XSSFCell.CELL_TYPE_ERROR) {
				System.out.println("ERROR Column");
				childCell.setCellValue(parentCell.getErrorCellString());
			} else if (parentCell.getCellType() == XSSFCell.CELL_TYPE_FORMULA) {

				System.out.println("FORMULA Column");
				switch (parentCell.getCachedFormulaResultType()) {
				case XSSFCell.CELL_TYPE_NUMERIC:
					childCell.setCellValue(parentCell.getNumericCellValue());
					break;
				case XSSFCell.CELL_TYPE_STRING:
					childCell.setCellValue(parentCell.getStringCellValue());
					break;
				}
			} else {
				childCell.setCellValue(parentCell.toString());
			}
		}
		// TODO Auto-generated method stub
	}

}
