package com.mobigen.iqa.common.util;

import java.io.File;
import java.io.FileOutputStream;

import java.text.SimpleDateFormat;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFPrintSetup;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;

public class ExcelCreateHelper {
	
	private String _fileName=null;
	private String _fullPath=null;
	private String _defaultSheetName = "data";
	private int _sheetCount = 0; 	
	private String[] _writeHeader = null;
	
	HSSFWorkbook _workbook = new HSSFWorkbook();	  
	HSSFSheet _sheet = null;
	HSSFRow _row = null;
	
	private String _title = null;	
	private int _writeRow = 0;
	
	public String get_DefaultSheetName() {
		return _defaultSheetName;
	}
	public void set_DefaultSheetName(String _DefaultSheetName) {
		this._defaultSheetName = _DefaultSheetName;
	}
	
	public String getFileNameMake(String menuName , String sLoginId){
		SimpleDateFormat formatter_ymdhms = new SimpleDateFormat ("yyyyMMddHHmmss");			
		java.util.Date curTime = new java.util.Date();
		String file_time = formatter_ymdhms.format(curTime);
		String fileName = menuName +"_"+sLoginId + "_" + file_time + ".xlsx";			
		return fileName;
	}
	
	public void init(String fullPath,String filename)
	{
		_fullPath = fullPath;
		_fileName = filename;
		
		_writeHeader = null;
		
		_writeRow = 0;
		_sheetCount = 0;
		
		try{
			_workbook = new HSSFWorkbook();						
		}
		catch(Exception e)
		{
			 System.out.println(e.getMessage());
		}
		
	}
	
	public void sheet_add()
	{
		_sheet = _workbook.createSheet(_defaultSheetName + "_" + String.valueOf(_sheetCount + 1));		
		_sheetCount++;
		_writeRow = 0;		
	}
	
	public void headerSetting(String str)
	{
		_writeHeader = str.split("\\|\\^\\|");
	}
	
	
	public void titleSetting(String title)
	{
		_title = title;
	}
	
	public void titleHeaderToWrite()
	{
		_writeRow++;
		write_title(_title);
		write_header();
	}
	
	private void write_title(String title)
	{
		_title = title;

		if ((_title == null) || (_title.isEmpty() == true))
		{							
			return;
		}
		/*
		try{
			
			HSSFFont font = _workbook.createFont();
			font.setFontName(HSSFFont.FONT_ARIAL);
			
			HSSFCellStyle titlestyle = _workbook.createCellStyle();
			titlestyle.setFillForegroundColor(HSSFColor.SKY_BLUE.index);
			titlestyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
			titlestyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
			titlestyle.setFont(font);
			
			HSSFCell cell = _row.createCell(0);
			cell.setCellValue(_title);
			cell1.setCellStyle(titlestyle);

						
			_writeRow++;
		}
		catch(Exception e)
		{
			
		}
		*/
		
	}
	
	private void write_header()
	{		
		try{
			
			sheet_add();
			
			HSSFFont font = _workbook.createFont();
			font.setFontName(HSSFFont.FONT_ARIAL);
			
			HSSFCellStyle titlestyle = _workbook.createCellStyle();
			titlestyle.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
			titlestyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
			titlestyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
			titlestyle.setFont(font);
			
			if (_writeHeader != null){
				if (_writeHeader.length > 0){
					
					HSSFCell cell = null;
					_row  = _sheet.createRow(_writeRow);
					for (int i=0;i<_writeHeader.length;i++)
					{
						cell = _row.createCell(i);
						cell.setCellValue(_writeHeader[i]);
						cell.setCellStyle(titlestyle);						
					}
					
					_writeRow++;
				}
			}			
		}catch(Exception e)
		{
			
		}
	}
	
	public void write_data(String strData){
		
		String[] writeData = strData.split("\\|\\^\\|");
		
		try
		{
			
			if ((_sheet == null) || (_writeRow <= 0 )){
				sheet_add();
				write_header();
			}
			
			_row  = _sheet.createRow(_writeRow);
			HSSFCell cell = null;
			for (int i=0;i<writeData.length;i++)
			{
				String strTmp = writeData[i].toString();
				cell = _row.createCell(i);
				cell.setCellValue(strTmp);					
			}
			
			_writeRow++;
			
		}catch(Exception e)
		{
			
		}
		
	}
	
	public void excelFileMake(){
		
		try
		{			
			/*
			HSSFPrintSetup ps = _sheet.getPrintSetup();
			_sheet.setAutobreaks(true);			
			ps.setFitWidth((short)1);
			*/
			
			File file = new File(_fullPath + _fileName);
			FileOutputStream fileOutput = new FileOutputStream(file);
			
			_workbook.write(fileOutput);
			fileOutput.close();
			
		}catch(Exception e)
		{
			
		}finally{
			close();
		}	
	}
	
	private void close()
	{
		try
		{
			_row = null;
			_sheet = null;			
			_workbook = null;
			
		}
		catch(Exception e)
		{
			
		}
	}
	
	
}
