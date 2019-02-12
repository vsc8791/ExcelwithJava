package main;

import entities.*;

import java.io.File;
import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;

import dao.*;

public class Main {

	public static void main(String[] args) {
		// TODO Auto-generated method stub

		try {
			ProductModel pmodel = new ProductModel();
			HSSFWorkbook workbook = new HSSFWorkbook();
			HSSFSheet sheet = workbook.createSheet("List Products");
			// Create Heading

			HSSFRow rowHeading = sheet.createRow(0);
			rowHeading.createCell(0).setCellValue("Id");
			rowHeading.createCell(1).setCellValue("Name");
			rowHeading.createCell(2).setCellValue("Creation Date");
			rowHeading.createCell(3).setCellValue("Price");
			rowHeading.createCell(4).setCellValue("Quantity");
			rowHeading.createCell(5).setCellValue("Sub Total");

			for (int i = 0; i < 6; i++) {

				HSSFCellStyle stylerowHeading = workbook.createCellStyle();
				HSSFFont font = workbook.createFont();
				font.setBold(true);
				font.setFontName(HSSFFont.FONT_ARIAL);
				font.setFontHeightInPoints((short) 11);
				stylerowHeading.setFont(font);
				stylerowHeading.setVerticalAlignment(CellStyle.ALIGN_CENTER);
				rowHeading.getCell(i).setCellStyle(stylerowHeading);

			}

			
			int r=1;
			for(Product p:pmodel.findAll()) 
			{
				Row row=sheet.createRow(r);
				
				//Id Column
				Cell cellId=row.createCell(0);
				cellId.setCellValue(p.getId());
				//Id Name
				Cell cellName=row.createCell(1);
				cellName.setCellValue(p.getName());
				//Creation Date Column
				Cell cellCreationDate=row.createCell(2);
				cellCreationDate.setCellValue(p.getCreationDate());
				CellStyle styleCreationDate=workbook.createCellStyle();
				HSSFDataFormat df=workbook.createDataFormat();
				styleCreationDate.setDataFormat(df.getFormat("mm/dd/yyyy"));
				cellCreationDate.setCellStyle(styleCreationDate);
				
				//Price Column
				Cell cellPrice=row.createCell(3);
				cellPrice.setCellValue(p.getPrice());
				CellStyle stylePrice=workbook.createCellStyle();
				HSSFDataFormat cf=workbook.createDataFormat();
				stylePrice.setDataFormat(cf.getFormat("₹##,##0.00"));
				cellPrice.setCellStyle(stylePrice);
				//Quantity Column
				
				Cell cellQuantity=row.createCell(4);
				cellQuantity.setCellValue(p.getQuantity());
				
				//Sub Total Column
				
				Cell cellSubTotal=row.createCell(5);
				cellSubTotal.setCellValue(p.getQuantity()*p.getPrice());
				CellStyle styleSubTotal=workbook.createCellStyle();
				HSSFDataFormat sdf =workbook.createDataFormat();
				styleSubTotal.setDataFormat(sdf.getFormat("₹##,##0.00"));
				cellSubTotal.setCellStyle(styleSubTotal);
				//Total Column
				Row rowTotal=sheet.createRow(pmodel.findAll().size()+1);
				Cell cellTextTotal=rowTotal.createCell(0);
				cellTextTotal.setCellValue("Total");
			//	@SuppressWarnings("deprecation")
			//	CellRangeAddress region=CellRangeAddress.valueOf("A5:E5");
			//    sheet.addMergedRegion(region);
				CellStyle styleTotal=workbook.createCellStyle();
				HSSFFont fontTextTotal=workbook.createFont();
				fontTextTotal.setBold(true);
                fontTextTotal.setFontHeightInPoints((short)11);
                fontTextTotal.setColor(HSSFColor.GREEN.index);
                
                styleTotal.setFont(fontTextTotal);
                styleTotal.setVerticalAlignment(CellStyle.ALIGN_RIGHT);
                cellTextTotal.setCellStyle(styleTotal);
                
                //Total Value Column
                
                Cell cellTotalValue =rowTotal.createCell(5);
                cellTotalValue.setCellFormula("sum(F2:F5)");
                HSSFDataFormat dfTotalValue =workbook.createDataFormat();
                CellStyle styleTotalValue=workbook.createCellStyle();
                styleTotalValue.setDataFormat(dfTotalValue.getFormat("₹##,##0.00"));
                cellTotalValue.setCellStyle(styleTotalValue);
                
                
				//Iterating rows				
				r++;
			}
			// Autofit
						for (int i = 0; i < 6; i++) {
							sheet.autoSizeColumn(i);
						}
			

			// Save to Excel File
			FileOutputStream out = new FileOutputStream(new File("D:\\listProducts.xls"));
			workbook.write(out);
			out.close();
			workbook.close();
			System.out.println("Excel Written SuccessFully...");
		} catch (Exception e) {
			// TODO: handle exception
			System.out.println(e.getMessage());
		}

	}

}
