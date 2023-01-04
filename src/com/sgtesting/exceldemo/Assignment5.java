//Programatically write 20 Flower names and Colour Name into 1 and 2 column of First Sheet
package com.sgtesting.exceldemo;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class Assignment5 {

	public static void main(String[] args) {
		writeContent();

	}
	private static void writeContent()
	{
		FileOutputStream fout=null;
		Workbook wb=null;
		
		Sheet sh=null;
	    Row row=null;
	   
		Cell cell=null;
		
		try
		{
			wb=new XSSFWorkbook();
			sh=wb.createSheet("Sheet1");
			
			row=sh.createRow(0);
			cell=row.createCell(0);
			cell.setCellValue("Flower1");
			cell=row.createCell(1);
			cell.setCellValue("Colour1");
			
			row=sh.createRow(1);
			cell=row.createCell(0);
			cell.setCellValue("Flower2");
			cell=row.createCell(1);
			cell.setCellValue("Colour2");
			
			row=sh.createRow(2);
			cell=row.createCell(0);
			cell.setCellValue("Flower3");
			cell=row.createCell(1);
			cell.setCellValue("Colour3");
			
			row=sh.createRow(3);
			cell=row.createCell(0);
			cell.setCellValue("Flower4");
			cell=row.createCell(1);
			cell.setCellValue("Colour4");
			
			row=sh.createRow(4);
			cell=row.createCell(0);
			cell.setCellValue("Flower5");
			cell=row.createCell(1);
			cell.setCellValue("Colour5");
			
			row=sh.createRow(5);
			cell=row.createCell(0);
			cell.setCellValue("Flower6");
			cell=row.createCell(1);
			cell.setCellValue("Colour6");
			
			row=sh.createRow(6);
			cell=row.createCell(0);
			cell.setCellValue("Flower7");
			cell=row.createCell(1);
			cell.setCellValue("Colour7");
			
			row=sh.createRow(7);
			cell=row.createCell(0);
			cell.setCellValue("Flower8");
			cell=row.createCell(1);
			cell.setCellValue("Colour8");
			
			row=sh.createRow(8);
			cell=row.createCell(0);
			cell.setCellValue("Flower9");
			cell=row.createCell(1);
			cell.setCellValue("Colour9");
			
			row=sh.createRow(9);
			cell=row.createCell(0);
			cell.setCellValue("Flower10");
			cell=row.createCell(1);
			cell.setCellValue("Colour10");
			
			row=sh.createRow(10);
			cell=row.createCell(0);
			cell.setCellValue("Flower11");
			cell=row.createCell(1);
			cell.setCellValue("Colour11");
			
			row=sh.createRow(11);
			cell=row.createCell(0);
			cell.setCellValue("Flower12");
			cell=row.createCell(1);
			cell.setCellValue("Colour12");
			
			row=sh.createRow(12);
			cell=row.createCell(0);
			cell.setCellValue("Flower13");
			cell=row.createCell(1);
			cell.setCellValue("Colour13");
			
			row=sh.createRow(13);
			cell=row.createCell(0);
			cell.setCellValue("Flower14");
			cell=row.createCell(1);
			cell.setCellValue("Colour14");
			
			row=sh.createRow(14);
			cell=row.createCell(0);
			cell.setCellValue("Flower15");
			cell=row.createCell(1);
			cell.setCellValue("Colour15");
			
			row=sh.createRow(15);
			cell=row.createCell(0);
			cell.setCellValue("Flower16");
			cell=row.createCell(1);
			cell.setCellValue("Colour16");
			
			row=sh.createRow(16);
			cell=row.createCell(0);
			cell.setCellValue("Flower17");
			cell=row.createCell(1);
			cell.setCellValue("Colour17");
			
			row=sh.createRow(17);
			cell=row.createCell(0);
			cell.setCellValue("Flower18");
			cell=row.createCell(1);
			cell.setCellValue("Colour18");
			
			row=sh.createRow(18);
			cell=row.createCell(0);
			cell.setCellValue("Flower19");
			cell=row.createCell(1);
			cell.setCellValue("Colour19");
			
			row=sh.createRow(19);
			cell=row.createCell(0);
			cell.setCellValue("Flower20");
			cell=row.createCell(1);
			cell.setCellValue("Colour20");



			fout=new FileOutputStream("C:\\EXCEL\\Assignments Results\\Assignment5.xlsx");
			wb.write(fout);
		}catch (Exception e) 
		{
			e.printStackTrace();
		}
		finally
		{
			try
			{
				fout.close();
				wb.close();
			}catch (Exception e) 
			{
				e.printStackTrace();
			}
		}
	}

}
