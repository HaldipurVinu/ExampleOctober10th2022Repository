//Programatically write 20 vegetable names into 5th column of first sheet of excel sheet
package com.sgtesting.exceldemo;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class Assignment4_Alter {

	public static void main(String[] args) {
		VegetablesDemo();

	}
	private static  void VegetablesDemo()
	{
		FileOutputStream fout=null;
		Workbook wb=null;
		Sheet sh=null;
		Row row=null;
		Cell cell=null;

		try
		{
			wb=new XSSFWorkbook();
			sh=wb.createSheet("VegetablesDemo");

			for(int r=0;r<20;r++)
			{
				row=sh.createRow(r);
				cell=row.createCell(4);
				cell.setCellValue("Vegetable"+(r+1));
			}
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

			}catch (Exception e) 
			{
				e.printStackTrace();
			}
		}
	}
}
