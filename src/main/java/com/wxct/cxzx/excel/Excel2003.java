package com.wxct.cxzx.excel;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.util.ArrayList;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;



public class Excel2003 {	
	/**
	 * 读取整个sheet的数据
	 * @param file 文件名
	 * @param sheetName 
	 * @param rowNmb 行号,0开始
	 * @param startCell 起始列,0开始
	 * @author zhenr
	 * @throws IOException 
	 * */
	public static ArrayList<String[]> readSheet(String file,int sheetNumber) throws IOException{
		ArrayList<String[]> list=new ArrayList<String[]>();
		FileInputStream fileIn=new FileInputStream(file);
		HSSFWorkbook workbook = new HSSFWorkbook(fileIn);
		HSSFSheet sheet = workbook.getSheetAt(sheetNumber);			
		if(null!=sheet){
			int rowNum=sheet.getLastRowNum();
			for(int i=0;i<=rowNum;i++){
				HSSFRow row = sheet.getRow(i);
				String[] rowResult=null;
				if(null!=row){
					int columnNum=row.getLastCellNum();//列数是实际情况，作为数据index要减1
					rowResult=new String [columnNum];
					for(int j=0;j<columnNum;j++){
						HSSFCell cell=row.getCell(j);
						if(cell!=null){
							if(cell.getCellType()==0){
								DecimalFormat decimalFormat = new DecimalFormat("##0");//double格式化设置,不要科学计数法
								rowResult[j]=decimalFormat.format(cell.getNumericCellValue());
							}
							else if(cell.getCellType()==1){
								rowResult[j]=cell.getStringCellValue();
							}
						}
					}
					list.add(rowResult);
				}else{
					//logger.error("row不存在");
				}				
			}
		}else {
				//logger.error("sheet不存在");
		}
		fileIn.close();
		workbook.close();
		return list;
	}
	
	/**
	 * 读取一行
	 * 
	 * */
	public static String[] readRow(String file,int sheetNumber,int rowNmb) throws IOException{			
		FileInputStream fileIn=new FileInputStream(file);
		HSSFWorkbook workbook = new HSSFWorkbook(fileIn);
		HSSFSheet sheet = workbook.getSheetAt(sheetNumber);
		if(null!=sheet){
			HSSFRow row = sheet.getRow(rowNmb);
			if(null!=row){
				int columnNum=row.getLastCellNum();//列数是实际情况，作为数据index要减1
				String []rowResult=new String [columnNum];
				for(int j=0;j<columnNum;j++){
					HSSFCell cell=row.getCell(j);
					if(cell!=null){
						if(cell.getCellType()==0){
							DecimalFormat decimalFormat = new DecimalFormat("##0");//double格式化设置,不要科学计数法
							rowResult[j]=decimalFormat.format(cell.getNumericCellValue());
						}
						else if(cell.getCellType()==1){
							rowResult[j]=cell.getStringCellValue();
						}
					}							
				}
				workbook.close();
				return rowResult;
			}
		}
		workbook.close();
		return null;
	}
	
	/**
	 * 写入整个sheet的数据
	 * @param file 文件名
	 * @param sheetNumber 0开始sheet序号 
	 * @author zhenr
	 * @throws IOException 
	 * */
	public static void writeSheet(ArrayList<String[]> list,String file,int sheetNumber) throws IOException{		
		FileInputStream fileIn=new FileInputStream(file);
		HSSFWorkbook workbook = new HSSFWorkbook(fileIn);
		HSSFSheet sheet = workbook.getSheetAt(sheetNumber);
		//HSSFCellStyle style=workbook.createCellStyle();
		//style.setAlignment(HSSFCellStyle.ALIGN_CENTER);//居中
		if(null==sheet){
			sheet=workbook.createSheet();
		}
		int rowNum=0;
		for(String[] value : list){
			HSSFRow row = sheet.getRow(rowNum);
			if(null==row){
				row=sheet.createRow(rowNum);
			}
			for(int i=0;i<value.length;i++){
				if(null!=value[i]){
					String v=value[i];
					HSSFCell cell = row.getCell(i);
					if(null==cell){
						cell=row.createCell(i);
					}
					cell.setCellValue(v);
				}
				//if(null!=style)
					//cell.setCellStyle(style);
			}					
			rowNum++;
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
	public static void removeRow(String file,int sheetNumber,int rowNum) throws IOException{
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
	public static int getRowNum(String file,int sheetNumber) throws IOException{
		FileInputStream fileIn=new FileInputStream(file);
		HSSFWorkbook workbook = new HSSFWorkbook(fileIn);
		HSSFSheet sheet = workbook.getSheetAt(sheetNumber);
		if(null!=sheet){
			workbook.close();
			return sheet.getLastRowNum();
		}else {
			workbook.close();
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
				return row.getLastCellNum();
			}else {
				workbook.close();
				return 0;
			}
		}else {
			workbook.close();
			return 0;
		}
	}
///////////////////////////////////////////////////////////////////	容易异常
	/**
	 * 读取整个sheet的数据 带格式
	 * @param file 文件名
	 * @param sheetNumber 0开始sheet序号 
	 * @author zhenr
	 * @throws IOException 
	 * */
	public static ArrayList<HSSFCell[]> readSheetWithStyle(String file,int sheetNumber) throws IOException{
		ArrayList<HSSFCell[]> list=new ArrayList<HSSFCell[]>();				
		FileInputStream fileIn=new FileInputStream(file);
		HSSFWorkbook workbook = new HSSFWorkbook(fileIn);
		HSSFSheet sheet = workbook.getSheetAt(sheetNumber);			
		if(null!=sheet){
			int rowNum=sheet.getLastRowNum();
			for(int i=0;i<=rowNum;i++){
				HSSFRow row = sheet.getRow(i);
				HSSFCell[] rowResult=null;
				if(null!=row){
					int columnNum=row.getLastCellNum();//列数是实际情况，作为数据index要减1
					//Data.COLUMNNUM=columnNum;
					rowResult=new HSSFCell [columnNum];
					for(int j=0;j<columnNum;j++){
						rowResult[j]=row.getCell(j);
					}
					list.add(rowResult);
				}else{
					//logger.error("row不存在");
				}				
			}
		}else {
				//logger.error("sheet不存在");
		}
		fileIn.close();
		workbook.close();
		return list;
	}
	
	/**
	 * 写入整个sheet的数据，带格式复制，易超出虚拟机内存
	 * @param file 文件名
	 * @param sheetNumber 0开始sheet序号 
	 * @author zhenr
	 * @throws IOException 
	 * */
	public static void writeSheetWithStyle(ArrayList<HSSFCell[]> list,String file,int sheetNumber) throws IOException{		
		FileInputStream fileIn=new FileInputStream(file);
		HSSFWorkbook workbook = new HSSFWorkbook(fileIn);
		HSSFSheet sheet = workbook.getSheetAt(sheetNumber);
		//HSSFCellStyle style=workbook.createCellStyle();
		//style.setAlignment(HSSFCellStyle.ALIGN_CENTER);//居中
		if(null==sheet){
			sheet=workbook.createSheet();
		}
		int rowNum=0;
		for(HSSFCell[] value : list){
			System.out.println(rowNum);
			HSSFRow row = sheet.getRow(rowNum);
			if(null==row){
				row=sheet.createRow(rowNum);
			}
			for(int i=0;i<value.length;i++){
				if(null!=value[i]){
					HSSFCell cell = row.getCell(i);
					if(null==cell){
						cell=row.createCell(i);
					}
					if(value[i].getCellType()==0){
						DecimalFormat decimalFormat = new DecimalFormat("##0");//double格式化设置,不要科学计数法
						cell.setCellValue(decimalFormat.format(value[i].getNumericCellValue()));
					}
					else if(value[i].getCellType()==1){
						cell.setCellValue(value[i].getStringCellValue());
					}
					if(null!=value[i].getCellStyle())
						cell.getCellStyle().cloneStyleFrom(value[i].getCellStyle());
				}
			}					
			rowNum++;
        }				
		FileOutputStream fileOut = new FileOutputStream(file);
		workbook.write(fileOut);
		workbook.close();
		fileIn.close();
		fileOut.close();	
	}
}
