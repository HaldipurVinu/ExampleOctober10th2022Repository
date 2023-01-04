//Programatically write 20 vegetable names into 5th column of first sheet of excel sheet
package com.sgtesting.exceldemo;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class Assignment4 {

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
			cell=row.createCell(4);
			cell.setCellValue("VegetableName1");
			
			row=sh.createRow(1);
			cell=row.createCell(4);
			cell.setCellValue("VegetableName2");
			row=sh.createRow(2);
			cell=row.createCell(4);
			cell.setCellValue("VegetableName3");
			row=sh.createRow(3);
			cell=row.createCell(4);
			cell.setCellValue("VegetableName4");
			row=sh.createRow(4);
			cell=row.createCell(4);
			cell.setCellValue("VegetableName5");
			row=sh.createRow(5);
			cell=row.createCell(4);
			cell.setCellValue("VegetableName6");
			row=sh.createRow(6);
			cell=row.createCell(4);
			cell.setCellValue("VegetableName7");
			row=sh.createRow(7);
			cell=row.createCell(4);
			cell.setCellValue("VegetableName8");
			row=sh.createRow(8);
			cell=row.createCell(4);
			cell.setCellValue("VegetableName9");
			row=sh.createRow(9);
			cell=row.createCell(4);
			cell.setCellValue("VegetableName10");
			row=sh.createRow(10);
			cell=row.createCell(4);
			cell.setCellValue("VegetableName11");
			row=sh.createRow(11);
			cell=row.createCell(4);
			cell.setCellValue("VegetableName12");
			row=sh.createRow(12);
			cell=row.createCell(4);
			cell.setCellValue("VegetableName13");
			row=sh.createRow(13);
			cell=row.createCell(4);
			cell.setCellValue("VegetableName14");
			row=sh.createRow(14);
			cell=row.createCell(4);
			cell.setCellValue("VegetableName15");
			row=sh.createRow(15);
			cell=row.createCell(4);
			cell.setCellValue("VegetableName16");
			row=sh.createRow(16);
			cell=row.createCell(4);
			cell.setCellValue("VegetableName17");
			row=sh.createRow(17);
			cell=row.createCell(4);
			cell.setCellValue("VegetableName18");
			row=sh.createRow(18);
			cell=row.createCell(4);
			cell.setCellValue("VegetableName19");
			row=sh.createRow(19);
			cell=row.createCell(4);
			cell.setCellValue("VegetableName20");

			fout=new FileOutputStream("C:\\EXCEL\\Assignments Results\\Assignment4.xlsx");
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
