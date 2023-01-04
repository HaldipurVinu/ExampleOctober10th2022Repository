//Programatically write 20 Fruit Names into First Sheet,First Column of the Excel Sheet//
package com.sgtesting.exceldemo;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class Assignment1 {

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
			cell.setCellValue("FruitName1");
			
			row=sh.createRow(1);
			cell=row.createCell(0);
			cell.setCellValue("FruitName2");
			row=sh.createRow(2);
			cell=row.createCell(0);
			cell.setCellValue("FruitName3");
			row=sh.createRow(3);
			cell=row.createCell(0);
			cell.setCellValue("FruitName4");
			row=sh.createRow(4);
			cell=row.createCell(0);
			cell.setCellValue("FruitName5");
			row=sh.createRow(5);
			cell=row.createCell(0);
			cell.setCellValue("FruitName6");
			row=sh.createRow(6);
			cell=row.createCell(0);
			cell.setCellValue("FruitName7");
			row=sh.createRow(7);
			cell=row.createCell(0);
			cell.setCellValue("FruitName8");
			row=sh.createRow(8);
			cell=row.createCell(0);
			cell.setCellValue("FruitName9");
			row=sh.createRow(9);
			cell=row.createCell(0);
			cell.setCellValue("FruitName10");
			row=sh.createRow(10);
			cell=row.createCell(0);
			cell.setCellValue("FruitName11");
			row=sh.createRow(11);
			cell=row.createCell(0);
			cell.setCellValue("FruitName12");
			row=sh.createRow(12);
			cell=row.createCell(0);
			cell.setCellValue("FruitName13");
			row=sh.createRow(13);
			cell=row.createCell(0);
			cell.setCellValue("FruitName14");
			row=sh.createRow(14);
			cell=row.createCell(0);
			cell.setCellValue("FruitName15");
			row=sh.createRow(15);
			cell=row.createCell(0);
			cell.setCellValue("FruitName16");
			row=sh.createRow(16);
			cell=row.createCell(0);
			cell.setCellValue("FruitName17");
			row=sh.createRow(17);
			cell=row.createCell(0);
			cell.setCellValue("FruitName18");
			row=sh.createRow(18);
			cell=row.createCell(0);
			cell.setCellValue("FruitName19");
			row=sh.createRow(19);
			cell=row.createCell(0);
			cell.setCellValue("FruitName20");



			fout=new FileOutputStream("C:\\EXCEL\\Assignments Results\\Assignment1.xlsx");
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
