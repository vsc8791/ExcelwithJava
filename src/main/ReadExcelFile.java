package main;

import java.io.*;
import java.util.*;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;

import java.text.*;

public class ReadExcelFile {

	public static void main(String[] args) {

		try {
			SimpleDateFormat sdf = new SimpleDateFormat("MM/dd/yyyy");
			HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream("D:\\listProducts.xls"));
			HSSFSheet sheet = workbook.getSheetAt(0);
			Iterator<Row> rowIterator = sheet.iterator();
			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();
				Iterator<Cell> cellIterator = row.cellIterator();
				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();
					switch (cell.getCellType()) {
					case Cell.CELL_TYPE_BOOLEAN:
						System.out.print(cell.getBooleanCellValue() + "\t\t");
						break;
					case Cell.CELL_TYPE_NUMERIC:
						if (HSSFDateUtil.isCellDateFormatted(cell)) {
							Date date = HSSFDateUtil.getJavaDate(cell.getNumericCellValue());
							String dateFmt = cell.getCellStyle().getDataFormatString();
							System.out.print(dateFmt + ":" + sdf.format(date) + "\t\t");
						} else
							System.out.print(cell.getNumericCellValue() + "\t\t");
						break;
					case Cell.CELL_TYPE_STRING:
						System.out.print(cell.getStringCellValue() + "\t\t");
						break;
					case Cell.CELL_TYPE_FORMULA:
						FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
						
						System.out.println(evaluator.evaluate(cell).getNumberValue() + "\t\t");
						break;

					}

				}
				System.out.println("");
			}
			workbook.close();
		} catch (Exception e) {
			// TODO: handle exception
		}

	}
}
