//Programatically write 20 City Names into First Sheet Diagonally
package com.sgtesting.exceldemo;

import java.io.FileOutputStream;


import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Assignment3_Alter {

	public static void main(String[] args) {
		CityNamesDiagonalDemo();

	}
	private static void CityNamesDiagonalDemo()
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
    			for(int c=1;c<=r+1;c++)
    			{
    				cell=row.createCell(r);
    				cell.setCellValue("City"+c);
    			}
    			fout=new FileOutputStream("C:\\EXCEL\\Assignments Results\\Assignment3.xlsx");
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
