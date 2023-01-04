//Programatically write 20 Colour names into 10th and 11th row of First Sheet
package com.sgtesting.exceldemo;
import java.io.FileOutputStream;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class Assignment6 {

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
			cell.setCellValue("Flower1");
			cell=row.createCell(1);
			cell.setCellValue("Flower2");
			cell=row.createCell(2);
			cell.setCellValue("Flower3");
			cell=row.createCell(3);
			cell.setCellValue("Flower");
			cell=row.createCell(4);
			cell.setCellValue("Flower5");
			cell=row.createCell(5);
			cell.setCellValue("Flower6");
			cell=row.createCell(6);
			cell.setCellValue("Flower7");
			cell=row.createCell(7);
			cell.setCellValue("Flower8");
			cell=row.createCell(8);
			cell.setCellValue("Flower9");
			cell=row.createCell(9);
			cell.setCellValue("Flower10");
			cell=row.createCell(10);
			cell.setCellValue("Flower11");
			cell=row.createCell(11);
			cell.setCellValue("Flower12");
			cell=row.createCell(12);
			cell.setCellValue("Flower13");
			cell=row.createCell(13);
			cell.setCellValue("Flower14");
			cell=row.createCell(14);
			cell.setCellValue("Flower15");
			cell=row.createCell(15);
			cell.setCellValue("Flower16");
			cell=row.createCell(16);
			cell.setCellValue("Flower17");
			cell=row.createCell(17);
			cell.setCellValue("Flower18");
			cell=row.createCell(18);
			cell.setCellValue("Flower19");
			cell=row.createCell(19);
			cell.setCellValue("Flower20");
			cell=row.createCell(20);
			
            row=sh.createRow(10);
            
			cell=row.createCell(0);
			cell.setCellValue("Colour1");
			cell=row.createCell(1);
			cell.setCellValue("Colour2");
			cell=row.createCell(2);
			cell.setCellValue("Colour3");
			cell=row.createCell(3);
			cell.setCellValue("Colour");
			cell=row.createCell(4);
			cell.setCellValue("Colour");
			cell=row.createCell(5);
			cell.setCellValue("Colour6");
			cell=row.createCell(6);
			cell.setCellValue("Colour7");
			cell=row.createCell(7);
			cell.setCellValue("Colour8");
			cell=row.createCell(8);
			cell.setCellValue("Colour9");
			cell=row.createCell(9);
			cell.setCellValue("Colour10");
			cell=row.createCell(10);
			cell.setCellValue("Colour11");
			cell=row.createCell(11);
			cell.setCellValue("Colour12");
			cell=row.createCell(12);
			cell.setCellValue("Colour13");
			cell=row.createCell(13);
			cell.setCellValue("Colour14");
			cell=row.createCell(14);
			cell.setCellValue("Colour15");
			cell=row.createCell(15);
			cell.setCellValue("Colour16");
			cell=row.createCell(16);
			cell.setCellValue("Colour17");
			cell=row.createCell(17);
			cell.setCellValue("Colour18");
			cell=row.createCell(18);
			cell.setCellValue("Colour19");
			cell=row.createCell(19);
			cell.setCellValue("Colour20");
			cell=row.createCell(20);





			fout=new FileOutputStream("C:\\EXCEL\\Assignments Results\\Assignment6.xlsx");
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
