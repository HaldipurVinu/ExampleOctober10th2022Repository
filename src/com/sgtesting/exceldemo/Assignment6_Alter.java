//Programatically write 20 Flower and Colour names into 10th and 11th row of First Sheet
package com.sgtesting.exceldemo;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class Assignment6_Alter {

	public static void main(String[] args) {
		FlowerColorDemo();

	}
	private static void FlowerColorDemo()
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

			row=sh.createRow(10);
			for(int c=0;c<=20;c++)
			{
				cell=row.createCell(c);
				cell.setCellValue("Color"+(c+1));
			}

			row=sh.createRow(9);
			for(int c=0;c<=20;c++)
			{
				cell=row.createCell(c);
				cell.setCellValue("Flower"+(c+1));
			}

			fout=new FileOutputStream("C:\\EXCEL\\Assignments Results\\Assignment6.xlsx");
			wb.write(fout);

		}catch (Exception e) 
		{
			e.printStackTrace();
		}
		finally
		{
			try
			{

			}catch (Exception e) 
			{
				e.printStackTrace();
			}
		}
	}
}
