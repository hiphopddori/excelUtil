package com.mobigen.iqa.common.util;

import java.awt.Color;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
/*
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
*/

  
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFAnchor;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

//Cell Color
//http://stackoverflow.com/questions/10912578/apache-poi-xssfcolor-from-hex-code/10924483#10924483

public class ExcelCreateWrapper {
	
	private String _fileName=null;
	private String _fullPath=null;
	private String _sheetName = "data";
	private int _sheetCount = 0; 	
	private String[] _writeHeader = null;
	private List _headerRow = null; 
	
	//Sheet  _sheet = null;
	//Row _row = null;
	XSSFWorkbook _workbook = null;
	XSSFSheet _sheet = null;	
	XSSFRow _row = null;
	
	private String _title = null;	
	private int _writeRow = 0;
	
	private boolean rowMergeMode = false;
	private short _headerRowHeight = 430;
	
	public short get_headerRowHeight() {
		return _headerRowHeight;
	}

	public void set_headerRowHeight(short _headerRowHeight) {
		this._headerRowHeight = _headerRowHeight;
	}

	public void rowMergeMode(boolean bRowMerge){
		rowMergeMode = bRowMerge;
		_headerRow = new ArrayList();
	}
	
	/* 컬럼을 병합 한다.
	 * 
	*/
	public void colMergeAply(){
		/*
		int nMaxCol = ((String[])_headerRow.get(0)).length;
		int nMaxRow = _headerRow.size();
		
		String sRemHeader = "";
		
		for(int nRow = 0; nRow < nMaxRow ; nRow++){
			
			int nMergeStartCol = -1;
			int nMergeEndCol = -1;
			
			for (int nCol = 0; nCol < nMaxCol; nCol++){
								
				String sTmpHeader = ((String[])_headerRow.get(nRow))[nCol];
				if (nCol == 0){
					sRemHeader = sTmpHeader;
					nMergeStartCol = 0;
					nMergeEndCol = 0;
				}
				
				if (sRemHeader.equals(sTmpHeader)){										
					if (nCol > 0){
						nMergeEndCol = nCol;
					}					
				}else{
					
					if (nCol > 0){
						sRemHeader = sTmpHeader;
					}
					
					if (nMergeStartCol >=0 && nMergeStartCol!=nMergeEndCol){												
						cellMerge(nRow+1,nMergeStartCol,nRow+1,nMergeEndCol);																
					}	
					
					nMergeStartCol = nMergeEndCol + 1;
					nMergeEndCol = nMergeEndCol + 1;	
				}					
			}			
		}
		*/
		colMergeAply(0);
		
	}
	
	public void colMergeAply(int nStartRow){
		
		int nMaxCol = ((String[])_headerRow.get(0)).length;
		int nMaxRow = _headerRow.size();
		
		String sRemHeader = "";
		
		for(int nRow = nStartRow; nRow < nStartRow + nMaxRow ; nRow++){
			
			int nMergeStartCol = -1;
			int nMergeEndCol = -1;
			
			for (int nCol = 0; nCol < nMaxCol; nCol++){
								
				String sTmpHeader = ((String[])_headerRow.get(nRow-nStartRow))[nCol];
				if (nCol == 0){
					sRemHeader = sTmpHeader;
					nMergeStartCol = 0;
					nMergeEndCol = 0;
				}
				
				if (sRemHeader.equals(sTmpHeader)){										
					if (nCol > 0){
						nMergeEndCol = nCol;
					}					
				}else{
					
					if (nCol > 0){
						sRemHeader = sTmpHeader;
					}
					
					if (nMergeStartCol >=0 && nMergeStartCol!=nMergeEndCol){												
						cellMerge(nRow+1,nMergeStartCol,nRow+1,nMergeEndCol);																
					}	
					
					nMergeStartCol = nMergeEndCol + 1;
					nMergeEndCol = nMergeEndCol + 1;	
				}					
			}			
		}		
	}
	
	
	/* 행을 병합 한다.
	 * 
	*/
	public void rowMergeAply(){
		/*
		int nMaxCol = ((String[])_headerRow.get(0)).length;
		String sRemHeader = "";
		
		for(int nCol = 0; nCol < nMaxCol ; nCol++){
			
			int nMergeStartRow = -1;
			int nMergeEndRow = -1;
			
			for (int nRow = 0; nRow < _headerRow.size(); nRow++){
				String sTmpHeader =  "";
				
				sTmpHeader = ((String[])_headerRow.get(nRow))[nCol];
				if (nRow == 0){
					sRemHeader = sTmpHeader;
					nMergeStartRow = 0;
					nMergeEndRow = 0;
				}
				if (sRemHeader.equals(sTmpHeader)){										
					if (nRow > 0){
						nMergeEndRow = nRow;
					}					
				}else{
					
					if (nRow > 0){
						sRemHeader = sTmpHeader;
					}
					
					if (nMergeStartRow >=0 && nMergeStartRow!=nMergeEndRow){												
						cellMerge(nMergeStartRow+1, nCol, nMergeEndRow+1, nCol);																
					}	
					
					nMergeStartRow = nMergeEndRow + 1;
					nMergeEndRow = nMergeEndRow + 1;	
				}								
			}
			
			if (nMergeEndRow > 0 || (nMergeStartRow < nMergeEndRow)){
				//nMergeEndRow = _headerRow.size() -1;
				cellMerge(nMergeStartRow+1, nCol, nMergeEndRow+1, nCol);
			}						
		}	
		*/	
		rowMergeAply(0);
	}
	
	public void rowMergeAply(int nStartRow ){
		
		int nMaxCol = ((String[])_headerRow.get(0)).length;
		String sRemHeader = "";
		
		for(int nCol = 0; nCol < nMaxCol ; nCol++){
			
			int nMergeStartRow = -1;
			int nMergeEndRow = -1;
			
			for (int nRow = nStartRow; nRow < nStartRow + _headerRow.size(); nRow++){
				String sTmpHeader =  "";
				
				sTmpHeader = ((String[])_headerRow.get(nRow-nStartRow))[nCol];
				if (nRow == nStartRow){
					sRemHeader = sTmpHeader;
					nMergeStartRow = nStartRow;
					nMergeEndRow = nStartRow;
				}
				if (sRemHeader.equals(sTmpHeader)){										
					if (nRow > 0){
						nMergeEndRow = nRow;
					}					
				}else{
					
					if (nRow > 0){
						sRemHeader = sTmpHeader;
					}
					
					if (nMergeStartRow >=(0+nStartRow) && nMergeStartRow!=nMergeEndRow){												
						cellMerge(nMergeStartRow+1, nCol, nMergeEndRow+1, nCol);																
					}	
					
					nMergeStartRow = nMergeEndRow + 1;
					nMergeEndRow = nMergeEndRow + 1;	
				}								
			}
			
			if (nMergeEndRow > 0 || (nMergeStartRow < nMergeEndRow)){
				//nMergeEndRow = _headerRow.size() -1;
				cellMerge(nMergeStartRow+1, nCol, nMergeEndRow+1, nCol);
			}						
		}		
	}
	
	
	
	public void endRowMergeMode(){
		rowMergeMode = false;
		_headerRow.clear();
		_headerRow = null;
	}
	
	
	public void writeEntityList(List list , Object[] lstDataField) {
		writeEntityList(list,lstDataField,0);
	}
	
	/*  Remark : Hashmap으로 정의된 Data List를  엑셀로 출력한다.
	 *  Param  : list  -> DataList , lstDataField , Flex Header DataField 정보
	 *  
	 */
	public void writeHashMapList(List list , Object[] lstDataField , int nStartIndex) {
		
		 int nIndex = 0;
		
		 HashMap<String, Object> excelDataMap = new HashMap<String, Object>();
		
		 for(Object obj: list) {
			 
			 if (nIndex < nStartIndex){
	       		  nIndex++;
	       		  continue;
	       	  }
			 
			 excelDataMap = (HashMap)list.get(nIndex);
			 
			 String strData = "";
             for (int nCol = 0; nCol < lstDataField.length; nCol++){			
					
					String sDataField = lstDataField[nCol].toString();		
					String sDataPart = "";
					
					if (sDataField.length() > 0){				
						sDataPart = CUtil.null2str(excelDataMap.get(sDataField));						
					}else{
						sDataPart = "";
					}								
					strData+=sDataPart + "|^|";					
             }               
             write_data(strData);  
             
             nIndex++;
			 
		 }
		 
		 excelDataMap.clear();
         excelDataMap = null;
		
	}
	
	
	
	/*  Remark : Entity로 정의된 Data List를  엑셀로 출력한다.
	 *  Param  : list  -> DataList , lstDataField , Flex Header DataField 정보
	 *  
	 */
	public void writeEntityList(List list , Object[] lstDataField , int nStartIndex) {
		
          List<Map<String, Object>> resultList = new ArrayList<Map<String,Object>>();
         
          Map<String, Object> excelDataMap = new HashMap<String, Object>();
          
          int nIndex = 0;
          
      	  DecimalFormat formatDouble = new DecimalFormat();                    		
      	  formatDouble.setDecimalSeparatorAlwaysShown(false);  
      	  formatDouble.setGroupingUsed(false);
      	  
          for(Object obj: list) {
        	  
        	  if (nIndex < nStartIndex){
        		  nIndex++;
        		  continue;
        	  }

               java.lang.reflect.Field[] fields = obj.getClass().getDeclaredFields();              
               for(int i=0 ; i < fields.length ; i++ ) {

                    fields[i].setAccessible(true);
                    try {              
                    	
                    	if (fields[i].getType().toString().equals("class java.lang.Double")){
                    	     
                    		if (fields[i].get(obj) != null){
                    			excelDataMap.put(fields[i].getName(), formatDouble.format(fields[i].get(obj)));
                    		} else{                  			  
                			  excelDataMap.put(fields[i].getName(), "");
                		    }
                    		                    		
                    	}else{
                    		excelDataMap.put(fields[i].getName(), fields[i].get(obj));
                    	}
                    	                    	
                    } catch (IllegalArgumentException e) {
                         e.printStackTrace();
                    } catch (IllegalAccessException e) {
                         e.printStackTrace();
                    }
               } 
               
               String strData = "";
               for (int nCol = 0; nCol < lstDataField.length; nCol++){			
					
					String sDataField = lstDataField[nCol].toString();		
					String sDataPart = "";
					
					if (sDataField.length() > 0){				
						sDataPart = CUtil.null2str(excelDataMap.get(sDataField));						
					}else{
						sDataPart = "";
					}								
					strData+=sDataPart + "|^|";					
               }               
               write_data(strData);               
          }
          
          excelDataMap.clear();
          excelDataMap = null;
     }
	
	public void writeEntity(Object data , Object[] lstDataField) {
		
        List<Map<String, Object>> resultList = new ArrayList<Map<String,Object>>();
       
        Map<String, Object> excelDataMap = new HashMap<String, Object>();

        java.lang.reflect.Field[] fields = data.getClass().getDeclaredFields();         
        
        DecimalFormat formatDouble = new DecimalFormat();                    		
      	formatDouble.setDecimalSeparatorAlwaysShown(false);   
      	formatDouble.setGroupingUsed(false);
        
        for(int i=0 ; i < fields.length ; i++ ) {

              fields[i].setAccessible(true);
              try {                    	
            	  
            	  if (fields[i].getType().toString().equals("class java.lang.Double")){
            		  
            		  if (fields[i].get(data) != null){
            			  excelDataMap.put(fields[i].getName(), formatDouble.format(fields[i].get(data)));
            		  } else{
            			  excelDataMap.put(fields[i].getName(), "");
            		  }
            	  }else{
            		  excelDataMap.put(fields[i].getName(), fields[i].get(data));
            	  }
            	                    	
              } catch (IllegalArgumentException e) {
                   e.printStackTrace();
              } catch (IllegalAccessException e) {
                   e.printStackTrace();
              }
        }
         
        String sDataField = "";
        String strData = "";
        String sDataPart = "";
         for (int nCol = 0; nCol < lstDataField.length; nCol++){			
				
				sDataField = lstDataField[nCol].toString();		
								
				if (sDataField.length() > 0){				
					sDataPart = CUtil.null2str(excelDataMap.get(sDataField));						
				}else{
					sDataPart = "";
				}								
				strData+=sDataPart + "|^|";					
         }               
         write_data(strData);               
         
         
        excelDataMap.clear();
        excelDataMap = null;
   }
	
	
	//http://poi.apache.org/spreadsheet/quick-guide.html#Images 참조함
	//http://thinktibits.blogspot.kr/2012/12/Excel-XLSX-Insert-PNG-JPG-Image-Java-POI-Example.html
	/*  Remark : Sheet에 이미지 삽입
	 *  Param  : bytes-> 이미지 ,  col1-> 컬럼 시작접 , row1->행 시작점 , col2 - > 컬럼 end , row2-> Row end
	 */
	public void addImage(byte[] bytes , int col1 , int row1 , int col2 , int row2){
		 /*
		 int pictureIdx = _workbook.addPicture(bytes,Workbook.PICTURE_TYPE_PNG);		 
		 CreationHelper helper = _workbook.getCreationHelper();
		 
		 //Creates the top-level drawing patriarch.
		 //멀티 출력시 한번 선언해서 해야함
		 Drawing drawing = _sheet.createDrawingPatriarch();
		 //XSSFDrawing drawing = _sheet.createDrawingPatriarch();
		 
		 //Create an anchor that is attached to the worksheet
		 ClientAnchor anchor = helper.createClientAnchor();
		 anchor.setAnchorType(ClientAnchor.MOVE_AND_RESIZE);
		 
		 //set top-left corner for the image		 
		 anchor.setCol1(col1);		 
		 anchor.setRow1(row1);
		 anchor.setCol2(col2);
		 anchor.setRow2(row2);
		 
		 //Creates a picture
		 Picture pict = drawing.createPicture(anchor, pictureIdx);
		 
		 //Reset the image to the original size
		 pict.resize();
		 
		 anchor = null;
		 pict = null;
		 helper = null;
		 */
		
		int pictureIdx = _workbook.addPicture(bytes,Workbook.PICTURE_TYPE_PNG);		 

		 //멀티 출력시 한번 선언해서 해야함
		 XSSFDrawing drawing = _sheet.createDrawingPatriarch();		 

		 XSSFClientAnchor anchor = new XSSFClientAnchor();		 
		 anchor.setAnchorType(ClientAnchor.MOVE_AND_RESIZE);
		 
		 //이미지 영역 조절		 
		 anchor.setCol1(col1);		 
		 anchor.setRow1(row1);
		 anchor.setCol2(col2);
		 anchor.setRow2(row2);
		 
		 //이미지 생성
		 XSSFPicture  pict = drawing.createPicture(anchor, pictureIdx);		 
		 //이미지 리사이징 처리
		 pict.resize();		 
		 
		 pict = null;
		 anchor = null;
		 drawing = null;
		 		
	}
	
	public String getFileNameMake(String menuName , String sLoginId){
		SimpleDateFormat formatter_ymdhms = new SimpleDateFormat ("yyyyMMddHHmmss");			
		java.util.Date curTime = new java.util.Date();
		String file_time = formatter_ymdhms.format(curTime);
		String fileName = menuName +"_"+sLoginId + "_" + file_time + ".xlsx";			
		return fileName;
	}
	
	public String getSheetName() {
		return _sheetName;
	}

	public void setSheetName(String _sheetName) {
		this._sheetName = _sheetName;
	}
	
	
	public String addColumnStr(String sData){
		return sData+="|^|";
	}
	
	/* REMARK : nLoop 반복 횟수만큼 엑셀 포맷 스트링 생성
	
	*/
	public String addColumnStr(String sData,int nLoop){
		String sColStr = "";
		for(int nCnt = 1;nCnt <=nLoop; nCnt++){
			sColStr+=sData+"|^|";
		}			
		return sColStr;
	}
	

	public void init(String fullPath,String filename)
	{
		_fullPath = fullPath;
		_fileName = filename;
		
		_writeHeader = null;
		
		_writeRow = 0;
		_sheetCount = 0;
		
		try{
			/*
			File file = new File(_fullPath + _fileName);
			
			InputStream ist = new FileInputStream(file); 			
			_workbook = new XSSFWorkbook(ist);
			*/
			_workbook = new XSSFWorkbook();
		}catch(Exception ex){
			 System.out.println(ex.getMessage());
		}		
		
	}
	
	public void sheet_add()
	{
		//_sheet = _workbook.createSheet(_defaultSheetName + "_" + String.valueOf(_sheetCount + 1));		
		_sheet = _workbook.createSheet(_sheetName);
		_sheetCount++;
		_writeRow = 0;		
	}
	
	public void headerSetting(String str)
	{
		_writeHeader = str.split("\\|\\^\\|");
		
		if (rowMergeMode){
			_headerRow.add(_writeHeader);
		}
	}
	
	
	public void titleSetting(String title)
	{
		_title = title;
	}
	
	public void titleHeaderToWrite(boolean bAddSheet, boolean bAutoMerge)
	{		
		//write_title(_title);
		write_header(bAddSheet,bAutoMerge);
		//_writeRow++;
	}
	
	public String get2Alphabet(int number) {
		String alphabet = "";
		int asciiCode = 0;
		int power = getPower(number, 0);
		int prePower = power - 1;
		int preInterval = getInterval(prePower);
		int interval = number - preInterval;
		while (power > 0) {
			asciiCode = (int) (interval / Math.pow(26, power - 1));
			alphabet = alphabet + get2Ascii(asciiCode);
			interval = interval - (int) Math.pow(26, power - 1) * asciiCode;
			power--;
		}
		return alphabet;
	}
	
	private char get2Ascii(int intColumn) {
		int intAscii = 65;	// A 부터
		intAscii += intColumn;
		return (char) intAscii;
	}
	
	private int getPower(int num, int power) {
		if (num == 0) return 1;
		if (num < getInterval(power)) return power;
		else return getPower(num, power + 1);
	}

	private int getInterval(int power) {
		return (int) (Math.pow(26, power + 1) - 26) / 25;
	}
	
	
	public void cellMerge(int nRowForm , int nColFrom, int nRowTo , int nColTo ){
		
		String sColFrom = get2Alphabet(nColFrom);
		String sColTo = get2Alphabet(nColTo);
		
		String sRowFrom = String.valueOf(nRowForm);
		String sRowTo = String.valueOf(nRowTo);
		
		_sheet.addMergedRegion(CellRangeAddress.valueOf(sColFrom + "" + sRowFrom + ":" + sColTo + "" + sRowTo));
		
	}
	
	private void write_title(String title)
	{
		_title = title;

		if ((_title == null) || (_title.isEmpty() == true))
		{							
			return;
		}		
	}
	
	private void write_header()
	{	
		write_header(false,false);
	}
	
	private void write_header(boolean bAddSheet, boolean bAutoMerge)
	{		
		try{
			
			if (bAddSheet){				
				sheet_add();
			}
			
			XSSFFont font = _workbook.createFont();
			font.setFontName("aria");	
			font.setBold(true);
			font.setFontHeightInPoints((short)12);
			
			CellStyle titlestyle = _workbook.createCellStyle();		
			titlestyle.setFillForegroundColor(IndexedColors.GREY_40_PERCENT.getIndex());
			titlestyle.setFillPattern(CellStyle.SOLID_FOREGROUND);			
					
			titlestyle.setBorderBottom((short)1);
			titlestyle.setBorderLeft((short)1);
			titlestyle.setBorderRight((short)1);
			titlestyle.setBorderTop((short)1);

			titlestyle.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			titlestyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			titlestyle.setFont(font);
			titlestyle.setWrapText(true);
			
			if (_writeHeader != null){
				if (_writeHeader.length > 0){
					int nRemIdx = 0;
					String sRemData = "";
					String sTmpData = "";
					Cell cell = null;
					_row  = _sheet.createRow(_writeRow);
					_row.setHeight(_headerRowHeight);
					
					for (int i=0;i<_writeHeader.length;i++)
					{						
						sTmpData = CUtil.null2str(_writeHeader[i]);
						
						cell = _row.createCell(i);						
						cell.setCellValue(sTmpData);
						
						//if (sTmpData.indexOf("\r") >=0){
							cell.setCellStyle(titlestyle);
						//}else{
						//	cell.setCellStyle(titlestyle);
						//}
						
						if (bAutoMerge){
							if (i == 0){
								sRemData = sTmpData;							
							}						
							if (!sRemData.equals(sTmpData)){									
								if (nRemIdx != (i-1)){
									cellMerge(_writeRow+1,nRemIdx,_writeRow+1,i-1);																											
								}	
								nRemIdx = i;
								sRemData = sTmpData;
							}
							//마지막열 체크
							if (i == _writeHeader.length-1){
								if (sRemData.equals(sTmpData)){
									cellMerge(_writeRow+1,nRemIdx,_writeRow+1,i);
								}
							}							
						}									
					}					
					_writeRow++;
				}
			}					
		}catch(Exception e)
		{
			
		}
	}
	
	public void add_row(){
		_writeRow++;
	}
	
	public int  get_row(){
		return _writeRow;
	}
	
	public void write_data(String strData, boolean bNumberChk){
		
		String[] writeData = strData.split("\\|\\^\\|");
		
		try
		{
			
			if ((_sheet == null) || (_writeRow <= 0 )){
				sheet_add();
				write_header();
			}
			
			
			//CellStyle cellStyleNumber = _workbook.createCellStyle();		
			//cellStyleNumber.setDataFormat(_workbook.createDataFormat().getFormat("000000.000"));
			
			_row  = _sheet.createRow(_writeRow);
			Cell cell = null;
			for (int i=0;i<writeData.length;i++)
			{
				String sTmp = CUtil.stringNullToStr(writeData[i]);
				cell = _row.createCell(i);
				
				if (bNumberChk && CUtil.isNumber(sTmp) && sTmp.length() > 0){
					cell.setCellType(Cell.CELL_TYPE_NUMERIC);
					cell.setCellValue(Double.valueOf(sTmp));
					//cell.setCellStyle(cellStyleNumber);
				}else{
					cell.setCellType(Cell.CELL_TYPE_STRING);
					cell.setCellValue(CUtil.stringNullToStr(writeData[i]));
				}				
			}
			
			_writeRow++;
			
		}catch(Exception e)
		{
			
		}	
		
	}
	
	public void write_data(String strData){		
		write_data(strData, true);		
	}
	
	//Cell _cell = null;		
	public void update_cellColor(int nRow , int nCol ,  short cellColor){
		_row  = _sheet.getRow(nRow);
		
		CellStyle cellStyle = _workbook.createCellStyle();					
		cellStyle.setFillForegroundColor(cellColor);
		cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);		
		cellStyle.setBorderBottom((short)1);
		cellStyle.setBorderLeft((short)1);
		cellStyle.setBorderRight((short)1);
		cellStyle.setBorderTop((short)1);				
		 _row.getCell(nCol).setCellStyle(cellStyle);		
		
	}
	
	/* cellColor Color로 지정 데이타 처리
	 * 
	*/
	public void write_data_cellColor(String strData,Color[] cellColor){
		
		String[] writeData = strData.split("\\|\\^\\|");
		
		try
		{			
			if ((_sheet == null) || (_writeRow <= 0 )){
				sheet_add();
				write_header();
			}
			
			XSSFFont font = _workbook.createFont();
			font.setFontName("aria");	
			font.setFontHeightInPoints((short)10);
			//font.setColor(color)
			
			
			_row  = _sheet.createRow(_writeRow);
			Cell cell = null;
			for (int i=0;i<writeData.length;i++)
			{			
				cell = _row.createCell(i);
				
				if (cellColor.length -1 >= i){
					
					if (cellColor[i] != null){
						XSSFCellStyle cellStyle = _workbook.createCellStyle();
						XSSFColor myColor = new XSSFColor(cellColor[i]); 
						cellStyle.setFillForegroundColor(myColor);
						cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);		
						cellStyle.setBorderBottom((short)1);
						cellStyle.setBorderLeft((short)1);
						cellStyle.setBorderRight((short)1);
						cellStyle.setBorderTop((short)1);	
						cellStyle.setFont(font);
						
						/* 엑셀 폰트 가독성이 좋아서 굳이 변경을 안하는거로 우선 하자.
						if (cellColor[i].getRed() <= 195 && cellColor[i].getGreen() < 100 && cellColor[i].getBlue() >=245){
							//폰트 흰색
						}else{
							//폰트 블랙
						}
						*/
						
						//myColor에 따라 폰트 칼라 바꾸자						
						cell.setCellStyle(cellStyle);						
					}					
				}
				
				String sTmp = CUtil.stringNullToStr(writeData[i]);
				
				if (CUtil.isNumber(sTmp) && sTmp.length() > 0){
					cell.setCellType(Cell.CELL_TYPE_NUMERIC);
					cell.setCellValue(Double.valueOf(sTmp));
				}else{
					cell.setCellType(Cell.CELL_TYPE_STRING);
					cell.setCellValue(CUtil.stringNullToStr(writeData[i]));
				}								
			}			
			_writeRow++;
			
		}catch(Exception e)
		{
			
		}		
	}
	
	public void write_data(String strData,short[] cellColor){
		
		String[] writeData = strData.split("\\|\\^\\|");
		
		//CellStyle cellStyle = _workbook.createCellStyle();		
		//cellStyle.setFillForegroundColor(IndexedColors.GREY_40_PERCENT.getIndex());
		//cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);			
		
		try
		{			
			if ((_sheet == null) || (_writeRow <= 0 )){
				sheet_add();
				write_header();
			}
			
			_row  = _sheet.createRow(_writeRow);
			Cell cell = null;
			for (int i=0;i<writeData.length;i++)
			{
			
				cell = _row.createCell(i);
				
				if (cellColor.length -1 >= i){
					
					if (cellColor[i] != -1){
						CellStyle cellStyle = _workbook.createCellStyle();
						cellStyle.setFillForegroundColor(cellColor[i]);
						cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);		
						cellStyle.setBorderBottom((short)1);
						cellStyle.setBorderLeft((short)1);
						cellStyle.setBorderRight((short)1);
						cellStyle.setBorderTop((short)1);							
						cell.setCellStyle(cellStyle);						
					}					
				}
				
				String sTmp = CUtil.stringNullToStr(writeData[i]);
				
				if (CUtil.isNumber(sTmp) && sTmp.length() > 0){
					cell.setCellType(Cell.CELL_TYPE_NUMERIC);
					cell.setCellValue(Double.valueOf(sTmp));
				}else{
					cell.setCellType(Cell.CELL_TYPE_STRING);
					cell.setCellValue(CUtil.stringNullToStr(writeData[i]));
				}				
				
			}
			
			_writeRow++;
			
		}catch(Exception e)
		{
			
		}		
	}
	
	
	public void excelFileMake(){
		
		try
		{						
			File file = new File(_fullPath + _fileName);
			FileOutputStream fileOutput = new FileOutputStream(file);
			
			_workbook.write(fileOutput);
			fileOutput.close();
			
		}catch(Exception e)			
		{
			System.out.println(e.getMessage());
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
