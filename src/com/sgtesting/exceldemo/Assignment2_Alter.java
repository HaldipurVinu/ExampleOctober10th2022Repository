//programatically write 20 flower names into 10th Row of First Sheet
package com.sgtesting.exceldemo;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class Assignment2_Alter {

	public static void main(String[] args) {
		FlowerNameDemo();

	}
	private static void FlowerNameDemo()
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

			row=sh.createRow(9);
			for(int c=0;c<20;c++)
			{
				cell=row.createCell(c);
				for(int i=1;i<=c+1;i++)
				{
					cell.setCellValue("Flower"+i);

					fout=new FileOutputStream("C:\\EXCEL\\Assignments Results\\Assignment2.xlsx");
					wb.write(fout);
				}
			}


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
