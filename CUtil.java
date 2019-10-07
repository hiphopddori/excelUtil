package com.mobigen.iqa.common.util;

import java.awt.Color;
import java.io.File;
import java.io.FileOutputStream;
import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Enumeration;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Pattern;
import java.util.regex.Matcher;

import javax.servlet.http.HttpServletRequest;

import org.codehaus.jettison.json.JSONArray;
import org.codehaus.jettison.json.JSONException;
import org.codehaus.jettison.json.JSONObject;
import javax.servlet.http.HttpServletRequest;


public class CUtil {

	public static String null2str(Object obj) {
		return null2str(obj,"");
	}

	public static int str2int(String data) {

		if( data == null ){
			return 0;
		}else{
			return Integer.parseInt(data);
		}
	}
	
	public static HashMap getHashMapRequestParam(HttpServletRequest request){
		
		HashMap param = new HashMap();		
		Enumeration enu = request.getAttributeNames();
		
		while (enu.hasMoreElements()) {
		   String key = (String)enu.nextElement() ;		   
		   param.put(key, request.getAttribute(key));
		 }
		
		return param;
	
	}
	
	public static String getJSVision(){
		
		 String sJsVision = DateUtil.getCurrentDate("yyyyMMddHHmmss");
		 return sJsVision;
	}
	
	/* List<Map> 객체를 Json 스트링으로 변환한다.
	 * 
	 */
	public String listmap_to_json_string(List<Map<String, Object>> list)
	{       
	    JSONArray json_arr=new JSONArray();
	    for (Map<String, Object> map : list) {
	        JSONObject json_obj=new JSONObject();
	        for (Map.Entry<String, Object> entry : map.entrySet()) {
	            String key = entry.getKey();
	            Object value = entry.getValue();
	            try {
	                json_obj.put(key,value);
	            } catch (JSONException e) {
	                // TODO Auto-generated catch block
	                e.printStackTrace();
	            }                           
	        }
	        json_arr.put(json_obj);
	    }
	    return json_arr.toString();
	}
	

	public static String stringNullToStr(Object obj){
		String sRet = null2str(obj);
		if (sRet.equals("null")){
			sRet = "";
		}
		return sRet;

	}

	public static String[] getSplit(int nArray, String sData , String sDelimiter){
		String[] ret = null;
		sData= CUtil.null2str(sData);

		if (sData.length() <=0){
			ret = new String[nArray];
		}else{
			ret = sData.split(sDelimiter);
		}
		return ret;
	}

	public static String null2str(Object obj , String sReplaceStr) {

		if( obj == null )
			return sReplaceStr;
		else{

			String sRet = obj.toString();

			if (sRet.length() <=0){
				sRet = sReplaceStr;
			}
			return sRet;
		}
	}

	public static boolean isNumber(String str){

		//Pattern p = Pattern.compile("([\\p{Digit}]+)(([.]?)([\\p{Digit}]+))?");
	    //Matcher m = p.matcher(str);
		//return m.matches();
		return str.matches("[-+]?\\d*\\.?\\d+");

		 //return str.replaceAll("[+-]?\\d+", "").equals("") ? true : false;
	}


	public static boolean getFileSecureCheck(String sFileName){

		boolean bChk = true;

		String sExtension = getFindExtensionName(sFileName);
		sExtension = sExtension.toUpperCase();

		if (sExtension.equals("JSP") || sExtension.equals("PHP") || sExtension.equals("ASP") || sExtension.equals("EXE")
				|| sExtension.equals("CLASS") || sExtension.equals("SH") || sExtension.equals("BAT") ){
			bChk = false;
		}

		return bChk;

	}

	public static String getFindExtensionName(final String path) {

		 String fullPath = path;

	     int firstIndex = 0;

	     while(fullPath.indexOf('.') != -1) {
	    	 firstIndex = fullPath.indexOf('.');
	         fullPath = fullPath.substring(firstIndex+1);
	     }

	     return fullPath;
	}



	/* REMARK : 파일 확장자로 가능한 스트링을 얻는다.
	 * PARAM  : sFileName - > 파일 확장자 입력값
	 * RETURN : String -> 입력가능 파일 확장자
	 */
	public static String getFileTitleReplace(String sFileName){

		sFileName =sFileName.replace("/", "");
		sFileName =sFileName.replace(":", "");
		sFileName =sFileName.replace("*", "");
		sFileName =sFileName.replace("?", "");
		sFileName =sFileName.replace("<", "");
		sFileName =sFileName.replace(">", "");
		sFileName =sFileName.replace("|", "");

		return sFileName;
	}

	public static double getStrToDouble(String sData){
		return 0;
	}

	public static double round(double d, int n) {
	      return Math.round(d * Math.pow(10, n)) / Math.pow(10, n);
	}

	public static double getRate(double d, int n){
		d = d * 100;
		double dRet = round(d,n);
		//dRet = dRet * 100;
		//double dRet = Math.round(d*100)/100.0;
		return dRet;
	}

	// 리스트 엔티티 값을 해시맵으로 변환한다.
	public static List<Map<String, Object>> getMapListToElementList(List list) {
          List<Map<String, Object>> resultList = new ArrayList<Map<String,Object>>();

          for(Object obj: list) {

               // Object의 변수
               java.lang.reflect.Field[] fields = obj.getClass().getDeclaredFields();

               Map<String, Object> map = new HashMap<String, Object>();

               for(int i=0 ; i < fields.length ; i++ ) {

                    // private 변수에 접근 허용
                    fields[i].setAccessible(true);
                    try {

                         // 변수 명을 key로 value 저장.
                         map.put(fields[i].getName(), fields[i].get(obj));
                    } catch (IllegalArgumentException e) {
                         e.printStackTrace();
                    } catch (IllegalAccessException e) {
                         e.printStackTrace();
                    }
               }
               resultList.add(map);
          }

          return resultList;
	}

	// 엔티티 값을 해시맵으로 변환한다.
	public static HashMap<String, Object> getEntityToMap(Object obj){
		HashMap<String, Object> ret = new HashMap<String, Object>();

		Field[] fields = obj.getClass().getDeclaredFields();

		for(int i=0; i < fields.length; i++){
			fields[i].setAccessible(true);

			try{
				ret.put(fields[i].getName(), fields[i].get(obj));
			}catch(IllegalArgumentException e){
				e.printStackTrace();
			}catch(IllegalAccessException e){
				e.printStackTrace();
			}
		}

		return ret;
	}

	/* hex 코드에 따른 rgb color 값을 얻는다.
	 * 
	 */
	public static Color hex2Rgb(String colorStr) {		
	    return new Color(
	            Integer.valueOf( colorStr.substring( 1, 3 ), 16 ),
	            Integer.valueOf( colorStr.substring( 3, 5 ), 16 ),
	            Integer.valueOf( colorStr.substring( 5, 7 ), 16 ) );
	}
	
	public static String Upload_ServerSave(String winf_Path, byte[] bytes, String fileName  , String login_id) throws Exception
	{

		SimpleDateFormat formatter_ymd = new SimpleDateFormat ("yyyyMMdd");
		SimpleDateFormat formatter_ymdhms = new SimpleDateFormat ("yyyyMMdd_HHmmss");

		java.util.Date curTime = new java.util.Date();
		String file_time = formatter_ymdhms.format(curTime);
		String dir_time = formatter_ymd.format(curTime);

		String path = winf_Path + dir_time;

		fileName = file_time + "_" + login_id + "_" + fileName;

		File tempDir = null;
		tempDir =  new File(path);
		tempDir.mkdirs();
		tempDir = null;

		fileName = path + File.separator + fileName;
		File f = new File(fileName);
		FileOutputStream fos = new FileOutputStream(f);
		fos.write(bytes);
		fos.close();

		return fileName;
	}
	
	
	public static String getClientIp(HttpServletRequest request){
		
		String ip = request.getHeader("X-FORWARDED-FOR"); 
        
        //proxy 환경일 경우
        if (ip == null || ip.length() == 0) {
            ip = request.getHeader("Proxy-Client-IP");
        }

        //웹로직 서버일 경우
        if (ip == null || ip.length() == 0) {
            ip = request.getHeader("WL-Proxy-Client-IP");
        }

        if (ip == null || ip.length() == 0) {
            ip = request.getRemoteAddr() ;
        }
		
		return ip;
	}
	
	
	
	
	
	
	
	
	
	
	
	
	
	


}
