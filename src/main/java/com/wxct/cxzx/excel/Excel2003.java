package com.wxct.cxzx.excel;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;



public class Excel2003 {
	/**
	 * 从excel文件中读取某个sheet的完整数据，根据sheet序号
	 * 读取内容存为Object，读取时可以 instanceof 判断是什么类型
	 * @param file excel文件路径；sheetIndex 0开始sheet序号
	 * @return ArrayList<ArrayList<String>> ArrayList嵌套ArrayList，实现二维数组_表结构，ArrayList可保留数据顺序，可以直接根据索引获取:get(0)
	 * @throws IOException 
	 * */
	public ArrayList<ArrayList<Object>> readSheet(String file,int sheetIndex) throws IOException {		
		ArrayList<ArrayList<Object>> sheetValue=new ArrayList<ArrayList<Object>>();
		FileInputStream fileIn=new FileInputStream(file);
		HSSFWorkbook workbook = new HSSFWorkbook(fileIn);
		HSSFSheet sheet = workbook.getSheetAt(sheetIndex);
		if(null==sheet) {
			workbook.close();
			fileIn.close();
			return null;
		}
		int rowNum=sheet.getLastRowNum();//rownum 从0开始
		for(int i=0;i<=rowNum;i++){
			ArrayList<Object> rowValue=new ArrayList<Object>();
			HSSFRow row = sheet.getRow(i);
			if(null!=row) {//行非空
				int columnNum=row.getLastCellNum();//columnNum 从1开始计数;而XSSFRow.getCell的索引是从0开始的
				for(int j=0;j<columnNum;j++){//开始读行
					HSSFCell cell=row.getCell(j);
					if(null!=cell) {//单元格非空
						switch (cell.getCellTypeEnum()) {//根据单元格类型读取
						case NUMERIC:
							rowValue.add(cell.getNumericCellValue());
							break;
						case STRING:
							rowValue.add(cell.getStringCellValue());
							break;
						case FORMULA://出现公式的时候
							rowValue.add(cell.getNumericCellValue());
							break;
						case BLANK:
							rowValue.add(null);
							break;
						default:
							rowValue.add(null);
							break;
						}
					}else {//空格，写入null
						rowValue.add(null);
					}
				}
				sheetValue.add(rowValue);
			}else {//空行，加入null
				sheetValue.add(null);
			}
			
		}
		workbook.close();
		fileIn.close();
		return sheetValue;
	}
	
	/**
	 * 从excel文件中读取某个sheet的完整数据，根据sheetName
	 * 读取内容存为Object，读取时可以 instanceof 判断是什么类型
	 * @param file excel文件路径；sheetIndex 0开始sheet序号
	 * @return ArrayList<ArrayList<String>> ArrayList嵌套ArrayList，实现二维数组_表结构，ArrayList可保留数据顺序，可以直接根据索引获取:get(0)
	 * @throws IOException 
	 * */
	public ArrayList<ArrayList<Object>> getSheet(String file,String sheetName) throws IOException {
		ArrayList<ArrayList<Object>> sheetValue=new ArrayList<ArrayList<Object>>();
		FileInputStream fileIn=new FileInputStream(file);
		HSSFWorkbook workbook = new HSSFWorkbook(fileIn);
		HSSFSheet sheet = workbook.getSheet(sheetName);
		if(null==sheet) {
			workbook.close();
			fileIn.close();
			return null;
		}
		int rowNum=sheet.getLastRowNum();//rownum 从0开始
		for(int i=0;i<=rowNum;i++){
			ArrayList<Object> rowValue=new ArrayList<Object>();
			HSSFRow row = sheet.getRow(i);
			if(null!=row) {//行非空
				int columnNum=row.getLastCellNum();//columnNum 从1开始计数;而XSSFRow.getCell的索引是从0开始的
				for(int j=0;j<columnNum;j++){//开始读行
					HSSFCell cell=row.getCell(j);
					if(null!=cell) {//单元格非空
						switch (cell.getCellTypeEnum()) {//根据单元格类型读取
						case NUMERIC:
							rowValue.add(cell.getNumericCellValue());
							break;
						case STRING:
							rowValue.add(cell.getStringCellValue());
							break;
						case FORMULA://出现公式的时候
							rowValue.add(cell.getNumericCellValue());
							break;
						case BLANK:
							rowValue.add(null);
							break;
						default:
							rowValue.add(null);
							break;
						}
					}else {//空格，写入null
						rowValue.add(null);
					}
				}
				sheetValue.add(rowValue);
			}else {//空行，加入null
				sheetValue.add(null);
			}
			
		}
		workbook.close();
		fileIn.close();
		return sheetValue;
	}
	
	/**
	 * 写入数据到一个sheet，从指定起始行、起始列开始写，并使用起始行格式写入
	 * @param list 需写入的数据，使用ArrayList类型以保持顺序；file目标文件，sheetNumber 指定sheet号，rowNum 起始行号 0开始，colmunNum 起始列号 0开始
	 * @throws IOException 
	 * */
	public void writeSheet(ArrayList<ArrayList<Object>> list,String file,int sheetNumber,int rowNum,int colmunNum) throws IOException {
		FileInputStream fileIn=new FileInputStream(file);
		HSSFWorkbook workbook = new HSSFWorkbook(fileIn);
		HSSFSheet sheet = workbook.getSheetAt(sheetNumber);
		if(null==sheet){
			fileIn.close();
			workbook.close();
			return;
		}
		for(ArrayList<Object> rowValue : list){
			if(rowValue!=null) {
				HSSFRow row = sheet.getRow(rowNum);
				if(null==row){
					row=sheet.createRow(rowNum);
				}
				int column=colmunNum;
				for(Object cellValue : rowValue) {
					if(null!=cellValue){
						HSSFCell cell = row.getCell(column);
						if(null==cell){
							cell=row.createCell(column);
						}
						if (cellValue instanceof Integer) {
							cell.setCellValue(((Integer) cellValue).intValue());
						}else if (cellValue instanceof Double) {
							cell.setCellValue(((Double) cellValue).doubleValue());
						}else if (cellValue instanceof String) {
							cell.setCellValue((String) cellValue);
						}
					}
					column++;
				}					
				rowNum++;
			}
        }
		FileOutputStream fileOut = new FileOutputStream(file);
		workbook.write(fileOut);
		workbook.close();
		fileIn.close();
		fileOut.close();
	}
	
	/**
	 * 删除某一行，rowNum是行号-1，2003版的shiftRows，startrow参数需要+1
	 * @param file
	 * @param sheetNumber
	 * @param rowNum
	 * @author zhenr
	 * */
	public void removeRow(String file,int sheetNumber,int rowNum) throws IOException{
		FileInputStream fileIn=new FileInputStream(file);
		HSSFWorkbook workbook = new HSSFWorkbook(fileIn);
		HSSFSheet sheet = workbook.getSheetAt(sheetNumber);
		if(null!=sheet){
			if(null!=sheet.getRow(rowNum))
				sheet.shiftRows(rowNum+1, sheet.getLastRowNum(), -1);
		}
		FileOutputStream fileOut = new FileOutputStream(file);
		workbook.write(fileOut);
		workbook.close();
		fileIn.close();
		fileOut.close();
	}
	
	/**
	 * 获取行数
	 * 
	 * */
	public int getRowNum(String file,int sheetNumber) throws IOException{
		FileInputStream fileIn=new FileInputStream(file);
		HSSFWorkbook workbook = new HSSFWorkbook(fileIn);
		HSSFSheet sheet = workbook.getSheetAt(sheetNumber);
		if(null!=sheet){
			workbook.close();
			fileIn.close();
			return sheet.getLastRowNum();
		}else {
			workbook.close();
			fileIn.close();
			return -1;
		}
	}
	
	/**
	 * 获取列数
	 * 
	 * */
	public static int getColumnNum(String file,int sheetNumber) throws IOException{
		FileInputStream fileIn=new FileInputStream(file);
		HSSFWorkbook workbook = new HSSFWorkbook(fileIn);
		HSSFSheet sheet = workbook.getSheetAt(sheetNumber);
		if(null!=sheet){
			HSSFRow row = sheet.getRow(0);
			if(null!=row){
				workbook.close();
				fileIn.close();
				return row.getLastCellNum();
			}else {
				workbook.close();
				fileIn.close();
				return -1;
			}
		}else {
			workbook.close();
			fileIn.close();
			return -1;
		}
	}

	/**
	 * 读取一行，带格式
	 * @throws IOException 
	 * */
	public ArrayList<HSSFCell> getRowWithStyle(String file,int sheetNumber,int rowNum) throws IOException{
		ArrayList<HSSFCell> resultRow=new ArrayList<HSSFCell>();
		FileInputStream fileIn=new FileInputStream(file);
		HSSFWorkbook workbook = new HSSFWorkbook(fileIn);
		HSSFSheet sheet = workbook.getSheetAt(sheetNumber);
		if(null==sheet){
			fileIn.close();
			workbook.close();
			return null;
		}
		HSSFRow row = sheet.getRow(rowNum);
		if(null!=row){
			int columnNum=row.getLastCellNum();//列数是实际情况，作为数据index要减1
			for(int j=0;j<columnNum;j++){
				HSSFCell cell=row.getCell(j);
				if(cell!=null) {
					resultRow.add(row.getCell(j));
				}else {
					resultRow.add(null);
				}
			}
		}else {
			fileIn.close();
			workbook.close();
			return null;
		}
		fileIn.close();
		workbook.close();
		return resultRow;
	}
	
	/**
	 * 设置startRow开始的行的Style为sampleRow的Style,startRowNo开始写的列
	 * @throws IOException 
	 * */
	public void setStyleFromRow(String file,int sheetNumber,int startRowNo,ArrayList<HSSFCell> styleRow) throws IOException {
		FileInputStream fileIn=new FileInputStream(file);
		HSSFWorkbook workbook = new HSSFWorkbook(fileIn);
		HSSFSheet sheet = workbook.getSheetAt(sheetNumber);
		if(null==sheet){
			fileIn.close();
			workbook.close();
			return;
		}		
		if(null!=styleRow&&!styleRow.isEmpty()) {
			int rowNum=sheet.getLastRowNum();
			for(int i=startRowNo;i<=rowNum;i++){
				HSSFRow row = sheet.getRow(i);
				if(null!=row){
					int columnNum=row.getLastCellNum();//列数是实际情况，作为数据index要减1					
					for(int j=0;j<columnNum;j++){
						HSSFCell cell = row.getCell(j);
						if(null==cell){
							cell=row.createCell(j);
						}
						if(null!=styleRow.get(j)&&null!=styleRow.get(j).getCellStyle()) {
							cell.setCellStyle(styleRow.get(j).getCellStyle());
						}
					}
				}				
			}
		}else {
			fileIn.close();
			workbook.close();
			return;
		}
		FileOutputStream fileOut = new FileOutputStream(file);
		workbook.write(fileOut);
		workbook.close();
		fileIn.close();
		fileOut.close();
	}
	
	/**
	 * 读取某一格内容
	 * @throws IOException 
	 * */
	public Object readCell(String file,int sheetNumber,int sampleRowNo,int startRowNo) throws IOException {
		Object result="";
		FileInputStream fileIn=new FileInputStream(file);
		HSSFWorkbook workbook = new HSSFWorkbook(fileIn);
		HSSFSheet sheet = workbook.getSheetAt(sheetNumber);
		if(null==sheet){
			fileIn.close();
			workbook.close();
			return null;
		}
		HSSFRow row = sheet.getRow(sampleRowNo);
		if(null==row){
			fileIn.close();
			workbook.close();
			return null;
		}
		HSSFCell cell=row.getCell(startRowNo);
		if(null==cell){
			fileIn.close();
			workbook.close();
			return null;
		}
		switch (cell.getCellTypeEnum()) {//根据单元格类型读取
		case NUMERIC:
			result=cell.getNumericCellValue();
			break;
		case STRING:
			result=cell.getStringCellValue();
			break;
		case FORMULA://出现公式的时候
			result=cell.getNumericCellValue();
			break;
		case BLANK:
			result=null;
			break;
		default:
			result=null;
			break;
		}
		fileIn.close();
		workbook.close();
		return result;
	}
	
	/**
	 * 带格式修改一格
	 * @throws IOException 
	 * */
	public void editCell(Object value,String file,int sheetNumber,int sampleRowNo,int startRowNo) throws IOException {
		FileInputStream fileIn=new FileInputStream(file);
		HSSFWorkbook workbook = new HSSFWorkbook(fileIn);
		HSSFSheet sheet = workbook.getSheetAt(sheetNumber);
		if(null==sheet){
			fileIn.close();
			workbook.close();
			return;
		}
		HSSFRow row = sheet.getRow(sampleRowNo);
		if(null==row){
			fileIn.close();
			workbook.close();
			return;
		}
		HSSFCell cell=row.getCell(startRowNo);
		if(null==cell){
			fileIn.close();
			workbook.close();
			return;
		}
		if (value instanceof Integer) {
			cell.setCellValue(((Integer) value).intValue());
		}else if (value instanceof Double) {
			cell.setCellValue(((Double) value).doubleValue());
		}else if (value instanceof String) {
			cell.setCellValue((String) value);
		}
		cell.setCellStyle(cell.getCellStyle());
		FileOutputStream fileOut = new FileOutputStream(file);
		workbook.write(fileOut);
		workbook.close();
		fileIn.close();
		fileOut.close();
	}
}
