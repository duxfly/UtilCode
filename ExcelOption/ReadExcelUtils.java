package com.jixiang.argo.union.utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.beanutils.BeanUtils;
import org.apache.commons.lang.ArrayUtils;
import org.apache.poi.util.ArrayUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.jixiang.argo.union.frame.RSBLL;
import com.jixiang.argo.union.tools.PoiExcelHelper;
import com.jx.service.union.entity.EmployEntity;

/***
 * Excel操作类
 * @author flower
 */
public class ReadExcelUtils {
	public List<EmployEntity> readXls() throws IOException {
		InputStream is = new FileInputStream("D:/aa.xlsx");
		XSSFWorkbook hssfWorkbook = new XSSFWorkbook(is);
		List<EmployEntity> aaa = new ArrayList<EmployEntity>();
		// 循环工作表Sheet
		for (int numSheet = 0; numSheet < hssfWorkbook.getNumberOfSheets(); numSheet++) {
			XSSFSheet hssfSheet = hssfWorkbook.getSheetAt(numSheet);
			if (hssfSheet == null) {
				continue;
			}
			// 循环行Row
			for (int rowNum = 1; rowNum <= hssfSheet.getLastRowNum(); rowNum++) {
				XSSFRow hssfRow = hssfSheet.getRow(rowNum);
				if (hssfRow != null) {
					EmployEntity a = new EmployEntity();
					XSSFCell username = hssfRow.getCell(0);
					XSSFCell phonenumber = hssfRow.getCell(1);
					XSSFCell address = hssfRow.getCell(2);
					a.setRealname(getValue(username));
					a.setPhonenumber(getValue(phonenumber));
					a.setAddress(getValue(address));
					aaa.add(a);
				}
			}
		}
		return aaa;
	}

	@SuppressWarnings("static-access")
	private String getValue(XSSFCell hssfCell) {
		if (hssfCell.getCellType() == hssfCell.CELL_TYPE_BOOLEAN) {
			// 返回布尔类型的值
			return String.valueOf(hssfCell.getBooleanCellValue());
		} else if (hssfCell.getCellType() == hssfCell.CELL_TYPE_NUMERIC) {
			// 返回数值类型的值
			return String.valueOf(hssfCell.getNumericCellValue());
		} else {
			// 返回字符串类型的值
			return String.valueOf(hssfCell.getStringCellValue());
		}
	}
	// 打印单元格数据  
    private static void printBody(ArrayList<ArrayList<String>> dataList) {  
        int index = 0;  
        for(ArrayList<String> data : dataList) {  
            index ++;  
            System.out.println();  
            for(String v : data) {  
                System.out.print("\t\t" + v);  
            }  
        }  
    } 
	public static void main(String[] args) throws Exception{
		String filePath = "D:/aa.xlsx";
		PoiExcelHelper helper = PoiExcelHelper.getPoiExcelHelper(filePath);  

        // 读取excel文件数据  
        ArrayList<ArrayList<String>> dataList = helper.readExcel(filePath, 0, "2-");
        ArrayList<ArrayList<String>> sheet_dataList = helper.readExcel(filePath, 1);
        Object[] tempSheet = new Object[]{}; 
        for(ArrayList<String> temp : sheet_dataList){
        	tempSheet = temp.toArray();
        }
        
        JSONArray jsonArray = new JSONArray();
        for(ArrayList<String> strData : dataList){
			JSONObject jsonObject = new JSONObject();
			for(int i = 0;i<strData.size();i++){
				jsonObject.put(tempSheet[i].toString(), strData.get(i));
			}
	        jsonArray.add(jsonObject);
        }
        System.out.println(jsonArray.toJSONString());
        List<EmployEntity> tempList = JSON.parseArray(jsonArray.toString(), EmployEntity.class);
        for(EmployEntity e : tempList){
        	RSBLL.getInstance().getEmployService().addNewEmployEntity(e);
        }
        // 打印单元格数据  
        //printBody(dataList);
	}
}
