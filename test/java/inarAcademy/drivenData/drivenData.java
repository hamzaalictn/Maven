package inarAcademy.drivenData;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.Scanner;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class drivenData {

	public static void main(String[] args) throws IOException {

		// create XSSFworkbook object + FileInputStream argument for finding the excel
		// file.
		Scanner input = new Scanner(System.in);
		String desiredData = input.next();
		FileInputStream fis = new FileInputStream("/Users/hamzaalicetin/Desktop/java/Book1.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);

		int numberOfSheet = workbook.getNumberOfSheets();

		String ElementLocation = findElementLocationFromExcelFile(workbook, numberOfSheet, desiredData);
		System.out.println(ElementLocation);
		

		input.close();
	}

	public static String findElementLocationFromExcelFile(XSSFWorkbook workbook, int numberOfSheet,
			String desiredData) {
		String ElementLocation = "";
		for (int i = 0; i < numberOfSheet; i++) {

			if (workbook.getSheetName(i).equalsIgnoreCase("TestCaseSheet")) {
				XSSFSheet sheet = workbook.getSheetAt(i);

				// Identify TestCases column by scanning the entire 1st row.

				boolean isFounded = true;
				String desiredResult = desiredData;
				int numberOfRow = 0;

				do {
					Iterator<Row> rows = sheet.iterator();

					while (isFounded) {
						Row CurrentRow = rows.next();
						Iterator<Cell> CurrentRowCells = CurrentRow.cellIterator();
						numberOfRow++;

						int columnIndex = 65;
						while (CurrentRowCells.hasNext()) {
							Cell value = CurrentRowCells.next();
							if (value.getStringCellValue().equalsIgnoreCase(desiredResult)) {

								char letterOfColumn = (char) columnIndex;
								System.out.println("Data : " + desiredResult);
								ElementLocation = letterOfColumn + " / " + numberOfRow;
								isFounded = false;
								break;
							}
							columnIndex++;
						}

					}

				} while (isFounded);

			}
		}
		return ElementLocation;
	}

}
