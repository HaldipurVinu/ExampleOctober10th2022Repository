//There is excel file , in Sheet1 it has 20 Country names and Capital names in 1st and 2nd column.
//Read the content and write it into 10th and 11th row 
package com.sgtesting.exceldemo;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Assignment8_Alter {

	public static void main(String[] args) {
		readContent();

	}
	
	private static void readContent()
	{
		FileInputStream fin=null;
		FileOutputStream fout=null;
		Workbook wb=null;
		Sheet sh1=null;
		Sheet sh2=null;
		Row rowsh1=null;
		Row rowsh2=null;
		Row rowsh22=null;
		Cell cellsh1=null;
		Cell cellsh2=null;
		try
		{
			fin=new FileInputStream("C:\\EXCEL\\Assignments Results\\Assignment8.xlsx");
			wb=new XSSFWorkbook(fin);
			
			sh1=wb.getSheet("Sheet1");
			sh2=wb.getSheet("Sheet2");
			if(sh2==null)
			{
				sh2=wb.createSheet("Sheet2");
			}
			int rc=sh1.getPhysicalNumberOfRows();
			rowsh2=sh2.createRow(9);
			rowsh22=sh2.createRow(10);
			
			for(int r=0;r<rc;r++)
			{
				rowsh1=sh1.getRow(r);
				cellsh1=rowsh1.getCell(0);
				cellsh2=rowsh2.createCell(r);
				cellsh2.setCellValue(cellsh1.getStringCellValue());
				
				
				cellsh1=rowsh1.getCell(1);
				cellsh2=rowsh22.createCell(r);
				cellsh2.setCellValue(cellsh1.getStringCellValue());
			}
			
			fout=new FileOutputStream("C:\\EXCEL\\Assignments Results\\Assignment8.xlsx");
			wb.write(fout);
		}catch (Exception e) {
			e.printStackTrace();
		}
		finally
		{
			try
			{
				fin.close();
				fout.close();
				wb.close();
			}catch (Exception e) {
				e.printStackTrace();
			}
		}
	}

}
