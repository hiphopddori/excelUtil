package com.mobigen.iqa.common.util;

import java.io.File;
import java.io.FileInputStream;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReaderUtil {
	
	private XSSFWorkbook workBook_xlxs = null;
	private HSSFWorkbook workBook_xls = null;
	private Boolean _ExcelType2007 = true;
	
	private CellStyle cellStyle= null; 
	private DataFormat format = null;

	
	public int getMaxRow(){
		
		int nMaxRow = 0;
		
		if (_ExcelType2007){
			nMaxRow = workBook_xlxs.getSheetAt(0).getLastRowNum();
		}else{
			nMaxRow = workBook_xls.getSheetAt(0).getLastRowNum();
		}	
		
		return nMaxRow + 1;
	}

	public void setWorkBook(HSSFWorkbook workBook) {
		this.workBook_xls = workBook;		
	}
	
	public void setWorkBook(XSSFWorkbook workBook) {
		this.workBook_xlxs = workBook;		
	}
	
	public ExcelReaderUtil(HSSFWorkbook wb){
		workBook_xls = wb;
		
	}
	
	public ExcelReaderUtil(XSSFWorkbook wb){
		workBook_xlxs = wb;		
	}
	
	public void close() {
		
		if (workBook_xlxs != null){
			workBook_xlxs = null;
		}
		
		if (workBook_xls != null){
			workBook_xls = null;
		}
		
		if (format != null){
			format = null;
		}			
	}
	
	@Override
	protected void finalize() throws Throwable{	
		close();				
	}
	
	
	public ExcelReaderUtil(){		
	}
	
	/* REMARK : 해당 경로의 엑셀 파일을 오픈한다.
	 * 
	*/
	public boolean setFileOpen(String sExcelFile){
		
		try{
			File file = new File(sExcelFile);			
			if (!file.exists() || !file.isFile() || !file.canRead()){
				return false;				
			}else{				
				
				if (file.getName().indexOf("xlsx") >=0){
					_ExcelType2007 = true;
				}else{
					_ExcelType2007 = false;
				}
				
				if (_ExcelType2007){
					workBook_xlxs = new XSSFWorkbook(new FileInputStream(file));
					cellStyle = workBook_xlxs.createCellStyle();
					format = workBook_xlxs.createDataFormat();
				}else{
					workBook_xls = new HSSFWorkbook(new FileInputStream(file));
					cellStyle = workBook_xls.createCellStyle();
					format = workBook_xls.createDataFormat();
					
				}
				
				return true;
			}	
		}catch(Exception e){		
			return false;			
		}			
	}
	
	/* REAMRK : 문자값을 리턴한다.
	 *          (엑셀 문서중 공백으로 들어가는것들이 많아서 추가함)
	 * 
	 */
	private String getCell_TYPE_STRING(String sData){
		
		String sRet = CUtil.null2str(sData);
		
		//sRet = sRet.trim();
		
		//포맷중 빈값일경우 "-" 로 넘어오는 경우 발생
		if (sRet.length() == 1){
			if (sRet.indexOf("-") >= 0){
				sRet = "0";
			}
		}
		
		return sRet;
		
	}
	
	private String getCELL_TYPE_NUMERIC(double data ,boolean bDot , String sFormat){
		
		String sRet = CUtil.null2str(data,"0");
		
  	    if (sRet.indexOf("E") >=0){
  		   double val = data;
      	   BigDecimal changeVal = new BigDecimal(val);
      	   sRet = changeVal.toString();
  	    }else{  		
  	    	
  	    	if (!bDot){		//소수점이 없을경우
  	    		sRet = String.valueOf(Math.round(data));
  	    	}
  	    
  	    }
  	   
  	    return sRet;
		
	}
	
	private String getCELL_TYPE_NUMERIC(double data ,boolean bDot){
		
		return getCELL_TYPE_NUMERIC(data,bDot,"");
		
	}
	
	
	public String getText(final Row row , final int colIdx , boolean bCellFormet , String sFormat){
		String sRet = getText(row,colIdx, false , bCellFormet,sFormat);
		return sRet;
	}
	
	
	public String getText(final Row row , final int colIdx){
		String sRet = getText(row,colIdx, false , true , "");
		return sRet;
	}
	
	
	
	public String getText(final Row row , final int colIdx , boolean bDot , boolean bCellFormet , String sFormat){
		
		String sRet = "";
		
		if (row == null){
			return sRet;
		}
		
		Cell cell = row.getCell(colIdx);		
		
		if (cell != null){			
			
			switch( cell.getCellType() ) {
			
				case Cell.CELL_TYPE_NUMERIC :
		           if(DateUtil.isCellDateFormatted(cell)) {						//DATE 형일경우		        	   
		        	   SimpleDateFormat formatter = new SimpleDateFormat("yyyy-MM-dd");
		        	   sRet = formatter.format(cell.getDateCellValue());		        	   		        		        	  
		           } else {		        	   
		        	   //cellStyle.setDataFormat(format.getFormat("###"));
		        	   if (bCellFormet){
		        		   
		        		   if (sFormat.length() > 0){		        			   
		        			   if (!sFormat.equals("DIST")){
		        				   cellStyle.setDataFormat(format.getFormat(sFormat));
		        				   cell.setCellStyle(cellStyle);
		        			   }		        			   
		        		   }else{
		        			   cellStyle.setDataFormat(format.getFormat("##0"));
		        			   cell.setCellStyle(cellStyle);
		        		   }		        		   			        	   		       	
		        	   }
		        	   
		        	   sRet = getCELL_TYPE_NUMERIC(cell.getNumericCellValue(),bDot,sFormat);		        	 	   
		           }
		           break;	
		        
				case Cell.CELL_TYPE_FORMULA :					
					try { 
				        RichTextString stringValue = cell.getRichStringCellValue(); 				        
				         sRet = stringValue.getString(); 
				   } catch (Exception e) { 
					   
					   if (bCellFormet){
		        		   cellStyle.setDataFormat(format.getFormat("##0"));
			        	   cell.setCellStyle(cellStyle);		        	
		        	   }	        	   		        	   
		        	   sRet = getCELL_TYPE_NUMERIC(cell.getNumericCellValue(),bDot);		        	     	   
				   } 
					
				   break;
				case Cell.CELL_TYPE_BLANK :
					sRet = "";
					break;
					
				default :
					sRet = CUtil.null2str(cell.getStringCellValue());								
					break;			
			}			
		}		
		return sRet; 
		
	}
	
	public String getText(final Row row , Cell cell , final int colIdx){
		
		String sRet = "";
		
		if (cell != null){
			cell = row.getCell(colIdx);		
			sRet = cell.getStringCellValue();
		}							
		return  sRet;		
	}
	
	public Row getRow(final int sheetIdx, final int rowIdx){
		
		if (_ExcelType2007){
			return workBook_xlxs.getSheetAt(sheetIdx).getRow(rowIdx);
		}else{
			return workBook_xls.getSheetAt(sheetIdx).getRow(rowIdx);
		}
		
	}
	
	public String getText(final int sheetIdx, final int rowIdx, final int colIdx){
		
		String sRet = "";
		
		if (_ExcelType2007){
			if (workBook_xlxs != null){
				XSSFSheet sheet = workBook_xlxs.getSheetAt(sheetIdx);
				
				Row row = sheet.getRow(rowIdx);
				Cell cell = row.getCell(colIdx);			
				sRet = cell.getStringCellValue();			
			}
		}else{
			
			if (workBook_xls != null){
				HSSFSheet sheet = workBook_xls.getSheetAt(sheetIdx);
				
				Row row = sheet.getRow(rowIdx);
				Cell cell = row.getCell(colIdx);			
				sRet = cell.getStringCellValue();			
			}
			
		}		
		return  sRet;		
	}
	
	
	/* REMARK : 엑셀의 Row Count를 얻는다.
	 * PARMA  : sheetIdx -> 시트 Index
	 * RETURN : int -> MAxRowCount 
	 */
	public int getMaxRow(final int sheetIdx){
		//workBook.getSheetAt(sheetIdx).getLastRowNum();
		if (_ExcelType2007){
			return workBook_xlxs.getSheetAt(sheetIdx).getPhysicalNumberOfRows();
		}else{
			return workBook_xls.getSheetAt(sheetIdx).getPhysicalNumberOfRows();
		}
		
		
	}
	
	/* REMARK : 엑셀의 Sheet Count를 얻는다.
	 * PARMA  : sheetIdx -> 시트 Index
	 * RETURN : int -> MAxRowCount 
	 */
	public int getMaxSheet(){
		//workBook.getSheetAt(sheetIdx).getLastRowNum();
		if (_ExcelType2007){
			return workBook_xlxs.getNumberOfSheets();
		}else{
			return workBook_xls.getNumberOfSheets();
		}
		
		
	}
	
	public int getMaxCol(final int sheetIdx){
		
		int nMaxCol = -1;
		
		if (_ExcelType2007){
			XSSFSheet sheet = workBook_xlxs.getSheetAt(sheetIdx);
			Row row = sheet.getRow(0);			
			if (row != null){
				nMaxCol =row.getLastCellNum();		//문제 있어 우선 추후 검토한다. - 중간에 빈공간 잇으면 거기를 맥스Col로 인식함
			}
		}else{
			HSSFSheet sheet = workBook_xls.getSheetAt(sheetIdx);
			Row row = sheet.getRow(0);			
			if (row != null){
				nMaxCol =row.getLastCellNum();		
			}
		}
		
		nMaxCol = 100;
		
		return nMaxCol;
		
	}
	
}
