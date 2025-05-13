package utilities;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelFileUtil {
Workbook wb;
//constructor for reading excel path
public ExcelFileUtil(String ExcelPath)throws Throwable
{
	FileInputStream fi = new FileInputStream(ExcelPath);
	wb= WorkbookFactory.create(fi);
}
//counting rows in a sheet
public int rowCount(String sheetname)
{
	return wb.getSheet(sheetname).getLastRowNum();
}
//reading cell data
public String getCellData(String sheetName,int row,int column) 
{
	String data="";
	if(wb.getSheet(sheetName).getRow(row).getCell(column).getCellType()==CellType.NUMERIC)
	{
		int celldata =(int)wb.getSheet(sheetName).getRow(row).getCell(column).getNumericCellValue();
		data =String.valueOf(celldata);
		
	}
	else
	{
		data =wb.getSheet(sheetName).getRow(row).getCell(column).getStringCellValue();
	}
	return data;
}
//method for writing results
public void setCellData(String sheetname,int row,int column,String Status,String WriteExcel)throws Throwable
{
	//get sheet from wb
	Sheet ws =wb.getSheet(sheetname);
	//get row from ws
	Row rowNum =ws.getRow(row);
	//create cell in row
	Cell cell = rowNum.createCell(column);
	//write status in cell
	cell.setCellValue(Status);
	if(Status.equalsIgnoreCase("Pass"))
	{
		CellStyle style = wb.createCellStyle();
		Font font = wb.createFont();
		font.setColor(IndexedColors.GREEN.getIndex());
		font.setBold(true);
		style.setFont(font);
		rowNum.getCell(column).setCellStyle(style);
	}
	else if(Status.equalsIgnoreCase("Fail"))
	{
		CellStyle style = wb.createCellStyle();
		Font font = wb.createFont();
		font.setColor(IndexedColors.RED.getIndex());
		font.setBold(true);
		style.setFont(font);
		rowNum.getCell(column).setCellStyle(style);
	}
	else if(Status.equalsIgnoreCase("Blocked"))
	{
		CellStyle style = wb.createCellStyle();
		Font font = wb.createFont();
		font.setColor(IndexedColors.BLUE.getIndex());
		font.setBold(true);
		style.setFont(font);
		rowNum.getCell(column).setCellStyle(style);
	}
	FileOutputStream fo = new FileOutputStream(WriteExcel);
	wb.write(fo);
}
}



















