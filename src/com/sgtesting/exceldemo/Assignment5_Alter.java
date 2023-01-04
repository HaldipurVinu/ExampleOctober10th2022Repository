//Programatically write 20 Flower names and Colour Name into 1 and 2 column of First Sheet
package com.sgtesting.exceldemo;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class Assignment5_Alter {

	public static void main(String[] args) {
		FlowerColorColumns();

	}
	 private static void FlowerColorColumns()
	    {
	    	FileOutputStream fout=null;
	    	Workbook wb=null;
	    	Sheet sh=null;
	    	Row row=null;
	    	Cell cell=null;
	    	
	    	try
	    	{
	    		wb=new XSSFWorkbook();
	    		sh=wb.createSheet("FlowerColor1");
	    		 
	    		for(int r=0;r<20;r++)
	    		{
	    			row=sh.createRow(r);
	    			cell=row.createCell(0);
	    			for(int i=1;i<=r+1;i++)
	    			{
	    				cell.setCellValue("flower"+i);
	    			}
	    			cell=row.createCell(1);
	    			for(int i=1;i<=r+1;i++)
	    			{
	    				cell.setCellValue("color"+i);
	    			}
	    			fout=new FileOutputStream("C:\\EXCEL\\Assignments Results\\Assignment5.xlsx");
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
