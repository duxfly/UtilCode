package com.jixiang.argo.union.tools;

import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/****
 * 操作2007版本的excel工具类 文件结尾为.xlsx
 * @author flower
 */
public class PoiExcel07Helper extends PoiExcelHelper{
	/** 获取sheet列表 */  
    public ArrayList<String> getSheetList(String filePath) {  
        ArrayList<String> sheetList = new ArrayList<String>(0);  
        try {  
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(filePath));  
            Iterator<Sheet> iterator = wb.iterator(); 
            while (iterator.hasNext()) {  
                sheetList.add(iterator.next().getSheetName());  
            }  
        } catch (Exception e) {  
            e.printStackTrace();  
        }  
        return sheetList;  
    }  
  
    /** 读取Excel文件内容 */  
    public ArrayList<ArrayList<String>> readExcel(String filePath, int sheetIndex, String rows, String columns) {  
        ArrayList<ArrayList<String>> dataList = new ArrayList<ArrayList<String>> ();  
        try {  
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(filePath));  
            XSSFSheet sheet = wb.getSheetAt(sheetIndex);  
              
            dataList = readExcel(sheet, rows, getColumnNumber(sheet, columns));  
        } catch (Exception e) {  
            e.printStackTrace();  
        }  
        return dataList;  
    }  
      
    /** 读取Excel文件内容 */  
    public ArrayList<ArrayList<String>> readExcel(String filePath, int sheetIndex, String rows, int[] cols) {  
        ArrayList<ArrayList<String>> dataList = new ArrayList<ArrayList<String>> ();  
        try {  
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(filePath));  
            XSSFSheet sheet = wb.getSheetAt(sheetIndex);  
              
            dataList = readExcel(sheet, rows, cols);  
        } catch (Exception e) {  
            e.printStackTrace();  
        }  
        return dataList;  
    } 
}
