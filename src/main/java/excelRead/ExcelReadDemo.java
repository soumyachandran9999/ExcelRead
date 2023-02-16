package excelRead;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReadDemo {

	public static void main(String[] args) throws IOException {
		String excelFilePathString = "C:\\Users\\user\\Desktop\\Soumya\\CoreJava\\ExcelRead\\Rbi.xlsx";// excel sheet
																										// file path
		FileInputStream inputStream = new FileInputStream(excelFilePathString);// class used to read the excel sheet
		XSSFWorkbook workbook = new XSSFWorkbook(inputStream);// class to handle the workbook of excel sheet
		XSSFSheet sheet = workbook.getSheetAt(0);// class to handle sheet inside workbook
		int rows = sheet.getLastRowNum();// to get number of rows
		int cols = sheet.getRow(1).getLastCellNum();// to get number of colums
		for (int i = 0; i <= rows; i++) {
			XSSFRow row = sheet.getRow(i);// class to handle row which took each row by using i value
			for (int c = 0; c < cols; c++) {
				XSSFCell cell = row.getCell(c);
				switch (cell.getCellType()) {
				case STRING:
					System.out.print(cell.getStringCellValue() + " \t ");
					break;
				case NUMERIC:
					System.out.print(cell.getNumericCellValue() + " \t ");
					break;
				case BOOLEAN:
					System.out.print(cell.getBooleanCellValue() + " \t ");
					break;
				default:
				}

			}
			System.out.println(" ");
		}

	}

}
