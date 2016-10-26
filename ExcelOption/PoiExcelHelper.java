package com.jixiang.argo.union.tools;

import java.util.ArrayList;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;

/****
 * 使用POI接口操作excel工具类
 * @author flower
 */
public abstract class PoiExcelHelper {
	
	/***根据文件名称获取Excel处理类****/  
    public static PoiExcelHelper getPoiExcelHelper(String filePath) {
        PoiExcelHelper helper;  
        if(filePath.indexOf(".xlsx")!=-1) {
            helper = new PoiExcel07Helper();  
        }else {  
            helper = new PoiExcel03Helper();  
        }  
        return helper;  
    }
    /***
     * 通过excel中的数据封装为JSON格式返回
     * @param dataList 要导入的excel数据
     * @param sheet_dataList 要导入的excel表列集合
     * @return 封装的arraylist对象字符串
     */
    public static String getExcelValueToArrayList(ArrayList<ArrayList<String>> dataList,ArrayList<ArrayList<String>> sheet_dataList){
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
        return jsonArray.toString();
    }
	
	/*****
	***** 行列范围参数中均采用“,”作为不连续值的分割符，采用“-”作为两个连续值的连接符，这样简化了用户的参数配置，同时也保留了配置的灵活性，例如： 
	*****（1）12-        表示查询范围为从第十二行(列)到EXCEL中有记录的最后一行(列)； 
    *****（2）12-24      表示查询范围为从第十二行(列)到第二十四行(列)； 
    *****（3）12-24，30  表示查询范围为从第十二行(列)到第二十四行(列)、第三十行(列)等；
	*********/
	public static final String SEPARATOR = ",";
	public static final String CONNECTOR = "-";

	/** 获取sheet名称列表，子类必须实现 */
	public abstract ArrayList<String> getSheetList(String filePath);

	/** 读取Excel文件数据[获取指定sheet的所有内容从第一行和第一列开始] */
	public ArrayList<ArrayList<String>> readExcel(String filePath, int sheetIndex) {
		return readExcel(filePath, sheetIndex, "1-", "1-");
	}

	/** 读取Excel文件数据[获取指定sheet中从第rows行和第一列开始] */
	public ArrayList<ArrayList<String>> readExcel(String filePath, int sheetIndex, String rows) {
		return readExcel(filePath, sheetIndex, rows, "1-");
	}

	/** 读取Excel文件数据 */
	public ArrayList<ArrayList<String>> readExcel(String filePath, int sheetIndex, String[] columns) {
		return readExcel(filePath, sheetIndex, "1-", columns);
	}

	/** 读取Excel文件数据，子类必须实现 */
	public abstract ArrayList<ArrayList<String>> readExcel(String filePath, int sheetIndex, String rows,
			String columns);

	/** 读取Excel文件数据 */
	public ArrayList<ArrayList<String>> readExcel(String filePath, int sheetIndex, String rows, String[] columns) {
		int[] cols = getColumnNumber(columns);

		return readExcel(filePath, sheetIndex, rows, cols);
	}

	/** 读取Excel文件数据，子类必须实现 */
	public abstract ArrayList<ArrayList<String>> readExcel(String filePath, int sheetIndex, String rows, int[] cols);

	/** 读取Excel文件内容 */
	protected ArrayList<ArrayList<String>> readExcel(Sheet sheet, String rows, int[] cols) {
		ArrayList<ArrayList<String>> dataList = new ArrayList<ArrayList<String>>();
		// 处理行信息，并逐行列块读取数据
		String[] rowList = rows.split(SEPARATOR);
		for (String rowStr : rowList) {
			if (rowStr.contains(CONNECTOR)) {
				String[] rowArr = rowStr.trim().split(CONNECTOR);
				int start = Integer.parseInt(rowArr[0]) - 1;
				int end;
				if (rowArr.length == 1) {
					end = sheet.getLastRowNum();
				} else {
					end = Integer.parseInt(rowArr[1].trim()) - 1;
				}
				dataList.addAll(getRowsValue(sheet, start, end, cols));
			} else {
				dataList.add(getRowValue(sheet, Integer.parseInt(rowStr) - 1, cols));
			}
		}
		return dataList;
	}

	/** 获取连续行、列数据 */
	protected ArrayList<ArrayList<String>> getRowsValue(Sheet sheet, int startRow, int endRow, int startCol,
			int endCol) {
		if (endRow < startRow || endCol < startCol) {
			return null;
		}

		ArrayList<ArrayList<String>> data = new ArrayList<ArrayList<String>>();
		for (int i = startRow; i <= endRow; i++) {
			data.add(getRowValue(sheet, i, startCol, endCol));
		}
		return data;
	}

	/** 获取连续行、不连续列数据 */
	private ArrayList<ArrayList<String>> getRowsValue(Sheet sheet, int startRow, int endRow, int[] cols) {
		if (endRow < startRow) {
			return null;
		}

		ArrayList<ArrayList<String>> data = new ArrayList<ArrayList<String>>();
		for (int i = startRow; i <= endRow; i++) {
			data.add(getRowValue(sheet, i, cols));
		}
		return data;
	}

	/** 获取行连续列数据 */
	private ArrayList<String> getRowValue(Sheet sheet, int rowIndex, int startCol, int endCol) {
		if (endCol < startCol) {
			return null;
		}

		Row row = sheet.getRow(rowIndex);
		ArrayList<String> rowData = new ArrayList<String>();
		for (int i = startCol; i <= endCol; i++) {
			rowData.add(getCellValue(row, i));
		}
		return rowData;
	}

	/** 获取行不连续列数据 */
	private ArrayList<String> getRowValue(Sheet sheet, int rowIndex, int[] cols) {
		Row row = sheet.getRow(rowIndex);
		ArrayList<String> rowData = new ArrayList<String>();
		for (int colIndex : cols) {
			rowData.add(getCellValue(row, colIndex));
		}
		return rowData;
	}

	/**
	 * 获取单元格内容
	 * 
	 * @param row
	 * @param column
	 *            a excel column string like 'A', 'C' or "AA".
	 * @return
	 */
	protected String getCellValue(Row row, String column) {
		return getCellValue(row, getColumnNumber(column));
	}

	/**
	 * 获取单元格内容
	 * 
	 * @param row
	 * @param col
	 *            a excel column index from 0 to 65535
	 * @return
	 */
	private String getCellValue(Row row, int col) {
		if (row == null) {
			return "";
		}
		Cell cell = row.getCell(col);
		return getCellValue(cell);
	}

	/**
	 * 获取单元格内容
	 * 
	 * @param cell
	 * @return
	 */
	private String getCellValue(Cell cell) {
		if (cell == null) {
			return "";
		}

		String value = cell.toString().trim();
		try {
			// This step is used to prevent Integer string being output with
			// '.0'.
			Float.parseFloat(value);
			value = value.replaceAll("\\.0$", "");
			value = value.replaceAll("\\.0+$", "");
			return value;
		} catch (NumberFormatException ex) {
			return value;
		}
	}

	/**
	 * Change excel column letter to integer number
	 * 
	 * @param columns
	 *            column letter of excel file, like A,B,AA,AB
	 * @return
	 */
	private int[] getColumnNumber(String[] columns) {
		int[] cols = new int[columns.length];
		for (int i = 0; i < columns.length; i++) {
			cols[i] = getColumnNumber(columns[i]);
		}
		return cols;
	}

	/**
	 * Change excel column letter to integer number
	 * 
	 * @param column
	 *            column letter of excel file, like A,B,AA,AB
	 * @return
	 */
	private int getColumnNumber(String column) {
		int length = column.length();
		short result = 0;
		for (int i = 0; i < length; i++) {
			char letter = column.toUpperCase().charAt(i);
			int value = letter - 'A' + 1;
			result += value * Math.pow(26, length - i - 1);
		}
		return result - 1;
	}

	/**
	 * Change excel column string to integer number array
	 * 
	 * @param sheet
	 *            excel sheet
	 * @param columns
	 *            column letter of excel file, like A,B,AA,AB
	 * @return
	 */
	protected int[] getColumnNumber(Sheet sheet, String columns) {
		// 拆分后的列为动态，采用List暂存
		ArrayList<Integer> result = new ArrayList<Integer>();
		String[] colList = columns.split(SEPARATOR);
		for (String colStr : colList) {
			if (colStr.contains(CONNECTOR)) {
				String[] colArr = colStr.trim().split(CONNECTOR);
				int start = Integer.parseInt(colArr[0]) - 1;
				int end;
				if (colArr.length == 1) {
					end = sheet.getRow(sheet.getFirstRowNum()).getLastCellNum() - 1;
				} else {
					end = Integer.parseInt(colArr[1].trim()) - 1;
				}
				for (int i = start; i <= end; i++) {
					result.add(i);
				}
			} else {
				result.add(Integer.parseInt(colStr) - 1);
			}
		}

		// 将List转换为数组
		int len = result.size();
		int[] cols = new int[len];
		for (int i = 0; i < len; i++) {
			cols[i] = result.get(i).intValue();
		}

		return cols;
	}
}
