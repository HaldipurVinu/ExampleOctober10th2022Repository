//programatically write 20 flower names into 10th Row of First Sheet
package com.sgtesting.exceldemo;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class Assignment2 {

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
            row=sh.createRow(9);
            
			cell=row.createCell(0);
			cell.setCellValue("FlowerName1");
			cell=row.createCell(1);
			cell.setCellValue("FlowerName2");
			cell=row.createCell(2);
			cell.setCellValue("FlowerName3");
			cell=row.createCell(3);
			cell.setCellValue("FlowerName4");
			cell=row.createCell(4);
			cell.setCellValue("FlowerName5");
			cell=row.createCell(5);
			cell.setCellValue("FlowerName6");
			cell=row.createCell(6);
			cell.setCellValue("FlowerName7");
			cell=row.createCell(7);
			cell.setCellValue("FlowerName8");
			cell=row.createCell(8);
			cell.setCellValue("FlowerName9");
			cell=row.createCell(9);
			cell.setCellValue("FlowerName10");
			
			cell=row.createCell(10);
			cell.setCellValue("FlowerName11");
			cell=row.createCell(11);
			cell.setCellValue("FlowerName12");
			cell=row.createCell(12);
			cell.setCellValue("FlowerName13");
			cell=row.createCell(13);
			cell.setCellValue("FlowerName14");
			cell=row.createCell(14);
			cell.setCellValue("FlowerName15");
			cell=row.createCell(15);
			cell.setCellValue("FlowerName16");
			cell=row.createCell(16);
			cell.setCellValue("FlowerName17");
			cell=row.createCell(17);
			cell.setCellValue("FlowerName18");
			cell=row.createCell(18);
			cell.setCellValue("FlowerName19");
			cell=row.createCell(19);
			cell.setCellValue("FlowerName20");
			cell=row.createCell(20);



			fout=new FileOutputStream("C:\\EXCEL\\Assignments Results\\Assignment2.xlsx");
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
