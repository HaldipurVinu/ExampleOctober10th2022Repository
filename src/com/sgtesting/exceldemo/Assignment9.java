//There is an excel file it has following column of data first name,last name,email,jobname,dept name,location 
//in this way 20 rows of data write a program to read the content from excel file and write it into a new excel file.   
package com.sgtesting.exceldemo;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class Assignment9 {

	public static void main(String[] args) {
		readWriteContent();

	}
	private static void readWriteContent()
	{
		FileInputStream fin=null;
		FileOutputStream fout=null;
		Workbook wb=null;
		Sheet sh1=null;
		
		Row rowSh1=null;
		
		Cell cellSh1=null;
		
		try
		{
			fin=new FileInputStream("C:\\EXCEL\\Assignments Results\\Assignment9.xlsx");
			wb=new XSSFWorkbook(fin);
			sh1=wb.getSheet("Sheet1");
			
					
			int rc=sh1.getPhysicalNumberOfRows();
			for(int r=0;r<rc;r++)
			{
				rowSh1=sh1.getRow(r);
				
				
				int cc=rowSh1.getPhysicalNumberOfCells();
				for(int c=0;c<cc;c++)
				{
					cellSh1=rowSh1.getCell(c);
					
					
					
				}
			}
			fout=new FileOutputStream("C:\\EXCEL\\Assignments Results\\Assignment9_New.xlsx");
			wb.write(fout);
		}catch (Exception e) 
		{
			e.printStackTrace();
		}
		finally
		{
			try
			{
				fin.close();
				fout.close();
				wb.close();
			}catch (Exception e) 
			{
				e.printStackTrace();
			}
		}
	}

}
