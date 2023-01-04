//There is an excel file it has 20 country name and capital name in 1st and 2nd column of First sheet
//Read the data and Write it to 9th and 10th column of Sheet 2
package com.sgtesting.exceldemo;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class Assignment7 {

	public static void main(String[] args) {
		readWriteContent();

	}
	private static void readWriteContent()
	{
		FileInputStream fin=null;
		FileOutputStream fout=null;
		Workbook wb=null;
		Sheet sh1=null;
		Sheet sh2=null;
		Row rowSh1=null;
		Row rowSh2=null;
		Cell cellSh1=null;
		Cell cellSh2=null;
		try
		{
			fin=new FileInputStream("C:\\EXCEL\\Assignments Results\\Assignment7.xlsx");
			wb=new XSSFWorkbook(fin);
			sh1=wb.getSheet("Sheet1");
			sh2=wb.getSheet("Sheet2");
			if(sh2==null)
			{
				sh2=wb.createSheet("Sheet2");
			}
					
			int rc=sh1.getPhysicalNumberOfRows();
			for(int r=0;r<rc;r++)
			{
				rowSh1=sh1.getRow(r);
				rowSh2=sh2.getRow(r);
				if(rowSh2==null)
				{
					rowSh2=sh2.createRow(r);
				}
				
				int cc=rowSh1.getPhysicalNumberOfCells();
				for(int c=0;c<cc;c++)
				{
					cellSh1=rowSh1.getCell(c);
					cellSh2=rowSh2.getCell(c+8);
					if(cellSh2==null)
					{
						cellSh2=rowSh2.createCell(c+8);
					}
					String data=cellSh1.getStringCellValue();
					cellSh2.setCellValue(data);
				}
			}
			fout=new FileOutputStream("C:\\EXCEL\\Assignments Results\\Assignment7.xlsx");
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
