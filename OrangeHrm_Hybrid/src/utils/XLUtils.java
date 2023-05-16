package utils;

import java.io.FileInputStream;
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

public class XLUtils 
{
	public static FileInputStream fi;
	public static FileOutputStream fo;
	public static Workbook wb;
	public static Sheet ws;
	public static Row row;
	public static Cell cell;
	public static CellStyle style;
	
	public static int getRowCount(String xlfile, String xlsheet) throws IOException
	
	{
		fi=new FileInputStream(xlfile);
		wb= new XSSFWorkbook(fi);
		ws=wb.getSheet(xlsheet);
		int row_count=ws.getLastRowNum();
		wb.close();
		return row_count;		
	}
	
	public static Short getColumnCount(String xlfile, String xlsheet, int rownum) throws IOException
	
	{
		fi=new FileInputStream(xlfile);
		wb= new XSSFWorkbook(fi);
		ws=wb.getSheet(xlsheet);
		row=ws.getRow(rownum);
		Short cell_count=row.getLastCellNum();
		wb.close();
		return cell_count;		
	}

	public static String getStringData(String xlfile,String xlsheet,int rownum,int cellnum) throws IOException
	{
		fi = new FileInputStream(xlfile);
		wb = new XSSFWorkbook(fi);
		ws = wb.getSheet(xlsheet);
		row = ws.getRow(rownum);
		cell = row.getCell(cellnum);
		String data;
		try 
		{
			data = cell.getStringCellValue();			
		} catch (Exception e) 
		{
			data = "";			
		}		
		wb.close();	
		return data;	
	}
	
	public static double getNumericData(String xlfile, String xlsheet,int rownum, int cellnum) throws IOException
	
	{
		fi=new FileInputStream(xlfile);
		wb= new XSSFWorkbook(fi);
		ws=wb.getSheet(xlsheet);
		row=ws.getRow(rownum);
		cell=row.getCell(cellnum);
		double data;
		try {
			
			data=cell.getNumericCellValue();
			
			} catch (Exception e) 
			{
				data=0.0;
				System.err.println("No data found !");
			}
		wb.close();
		return data;
	}
	
	
	public static boolean getBooleanData(String xlfile, String xlsheet,int rownum, int cellnum) throws IOException
	
	{
		fi=new FileInputStream(xlfile);
		wb= new XSSFWorkbook(fi);
		ws=wb.getSheet(xlsheet);
		row=ws.getRow(rownum);
		cell=row.getCell(cellnum);
		boolean data;
		try {
			
			data=cell.getBooleanCellValue();
			
			} catch (Exception e) 
			{
				data=false;
				System.err.println("No data found !");
			}
		wb.close();
		return data;
	}
	
	
	public static void setData(String xlfile, String xlsheet, int rownum, int cellnum, String data) throws IOException
	
	{
		fi=new FileInputStream(xlfile);
		wb= new XSSFWorkbook(fi);
		ws=wb.getSheet(xlsheet);
		row=ws.getRow(rownum);
		cell=row.createCell(cellnum);
		cell.setCellValue(data);
		
		fo=new FileOutputStream(xlfile);
		wb.write(fo);
		wb.close();		
	}
	
	public static void fillGreenColor(String xlfile, String xlsheet, int rownum, int cellnum) throws IOException
	
	{
		fi=new FileInputStream(xlfile);
		wb= new XSSFWorkbook(fi);
		ws=wb.getSheet(xlsheet);
		row=ws.getRow(rownum);
		cell=row.getCell(cellnum);
		
		style=wb.createCellStyle();
		style.setFillForegroundColor(IndexedColors.BRIGHT_GREEN.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		
		cell.setCellStyle(style);
		fo=new FileOutputStream(xlfile);
		wb.write(fo);
		wb.close();		
	}	
	
	
	public static void fillRedColor(String xlfile, String xlsheet, int rownum, int cellnum) throws IOException
	
	{
		fi=new FileInputStream(xlfile);
		wb= new XSSFWorkbook(fi);
		ws=wb.getSheet(xlsheet);
		row=ws.getRow(rownum);
		cell=row.getCell(cellnum);
		
		style=wb.createCellStyle();
		style.setFillForegroundColor(IndexedColors.RED.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		
		cell.setCellStyle(style);
		fo=new FileOutputStream(xlfile);
		wb.write(fo);
		wb.close();		
	}	

	
	
	
	
	
	
	
	
	
}
