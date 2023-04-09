package utils;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XLutils
{
	
	public static FileInputStream fi;
	public static FileOutputStream fo;
	public static Workbook wb;
	public static Sheet st;
	public static Row row;
	public static Cell cell;
	
	public static int getRowCount(String xlfile,String xlsheet) throws IOException
	{
		fi=new FileInputStream(xlfile); 
		wb=new XSSFWorkbook(fi);
		st=wb.getSheet(xlsheet);
		int row_count=st.getLastRowNum();
		wb.close();
		return row_count;	
	}
	public static int getColumnCount(String xlfile,String xlsheet,int rownum) throws IOException
	{
		fi=new FileInputStream(xlfile);
		wb=new XSSFWorkbook(fi);
		st=wb.getSheet(xlsheet);
		row=st.getRow(rownum);
		int col_count=row.getLastCellNum();
		wb.close();
		return col_count;
	}
	public static String getStringCellData(String xlfile,String xlsheet,int rownum,int colnum) throws IOException
	{
		fi=new FileInputStream(xlfile);
		wb=new XSSFWorkbook(fi);
		st=wb.getSheet(xlsheet);
		row=st.getRow(rownum);
		 
		String data;
		try {
			cell=row.getCell(colnum);
			data=cell.getStringCellValue();
			
		} catch (Exception e) {
			data="no data";
		}
		wb.close();
		return data;	
	}
	
	public static double getNumericCellData(String xlfile,String xlsheet,int rownum,int colnum) throws IOException
	{
		fi=new FileInputStream(xlfile);
		wb=new XSSFWorkbook(fi);
		st=wb.getSheet(xlsheet);
		row= st.getRow(rownum);
		
		double data;
		try {
		cell=row.getCell(colnum);
		data=cell.getNumericCellValue();
		}catch (Exception e) {
			data=0.0;
		}
				wb.close();
				return data;	
	}
	
	public static void setCellData(String xlfile,String xlsheet,int rownum,int colnum,String data) throws IOException
	{
		fi=new FileInputStream(xlfile);
		wb=new XSSFWorkbook(fi);
		st=wb.getSheet(xlsheet);
		row=st.getRow(rownum);
		cell=row.createCell(colnum);
		cell.setCellValue(data);
		fo=new FileOutputStream(xlfile);
		wb.write(fo);
		wb.close();	
	}	
	public static void fillGreenColor(String xlfile,String xlsheet,int rownum,int colnum,String xlfileres) throws IOException
	{
		fi=new FileInputStream(xlfile);
		wb=new XSSFWorkbook(fi);
		st=wb.getSheet(xlsheet);
		row=st.getRow(rownum);
		cell=row.getCell(colnum);
		CellStyle style=wb.createCellStyle();
		style.setFillForegroundColor(IndexedColors.BRIGHT_GREEN.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		cell.setCellStyle(style);
		fo=new FileOutputStream(xlfileres);
		wb.write(fo);
		wb.close();
	}
	public static void fillRedColor(String xlfile,String xlsheet,int rownum,int colnum,String xlfileres) throws IOException
	{
		fi=new FileInputStream(xlfile);
		wb=new XSSFWorkbook(fi);
		st=wb.getSheet(xlsheet);
		row=st.getRow(rownum);
		cell=row.getCell(colnum);
		CellStyle style=wb.createCellStyle();
		style.setFillForegroundColor(IndexedColors.RED.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		cell.setCellStyle(style);
		fo=new FileOutputStream(xlfileres);
		wb.write(fo);
		wb.close();
	}	
}
