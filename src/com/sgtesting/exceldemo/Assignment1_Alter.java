//Programatically write 20 Fruit Names into First Sheet,First Column of the Excel Sheet//
package com.sgtesting.exceldemo;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class Assignment1_Alter {

	public static void main(String[] args) {
		WriteFruitNames();

	}
	private static void WriteFruitNames()
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

			for(int r=0;r<20;r++)
			{
				row=sh.createRow(r);
				cell=row.createCell(0);
				for(int i=1;i<=r+1;i++)
				{
					for(int j=0;j<=i;j++)
					{
						cell.setCellValue("Fruit"+i); 

					}
				}
				fout=new FileOutputStream("C:\\EXCEL\\Assignments Results\\Assignment1.xlsx");
				wb.write(fout); 
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
