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
public class Assignment8 {

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
					cellSh2=rowSh2.getCell(c);
					if(cellSh2==null)
					{
						cellSh2=rowSh2.createCell(c);
					}
					String data=cellSh1.getStringCellValue();
					cellSh2.setCellValue(data);
				}
			}
			fout=new FileOutputStream("C:\\EXCEL\\Assignments Results\\Assignment8.xlsx");
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
