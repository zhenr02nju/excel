package com.wxct.cxzx.excel;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel2007 {
	/**
	 * 读取整个sheet的数据
	 * @param file 文件名
	 * @param sheetNumber 0开始sheet序号 
	 * @author zhenr
	 * @throws IOException 
	 * */
	public static ArrayList<String[]> readSheet(String file,int sheetNumber) throws IOException{
		ArrayList<String[]> list=new ArrayList<String[]>();				
		FileInputStream fileIn=new FileInputStream(file);
		XSSFWorkbook workbook = new XSSFWorkbook(fileIn);
		XSSFSheet sheet = workbook.getSheetAt(sheetNumber);			
		if(null!=sheet){
			int rowNum=sheet.getLastRowNum();
			for(int i=0;i<=rowNum;i++){
				XSSFRow row = sheet.getRow(i);
				String[] rowResult=null;
				if(null!=row){
					int columnNum=row.getLastCellNum();//列数是实际情况，作为数据index要减1
					//Data.COLUMNNUM=columnNum;
					rowResult=new String [columnNum];
					for(int j=0;j<columnNum;j++){
						XSSFCell cell=row.getCell(j);
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
		XSSFWorkbook workbook = new XSSFWorkbook(fileIn);
		XSSFSheet sheet = workbook.getSheetAt(sheetNumber);
		if(null!=sheet){
			XSSFRow row = sheet.getRow(rowNmb);
			if(null!=row){
				int columnNum=row.getLastCellNum();//列数是实际情况，作为数据index要减1
				String []rowResult=new String [columnNum];
				for(int j=0;j<columnNum;j++){
					XSSFCell cell=row.getCell(j);
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
	public static void writeSheet(List<String[]> list,String file,int sheetNumber) throws IOException{		
		FileInputStream fileIn=new FileInputStream(file);
		XSSFWorkbook workbook = new XSSFWorkbook(fileIn);
		XSSFSheet sheet = workbook.getSheetAt(sheetNumber);
		//XSSFCellStyle style=workbook.createCellStyle();
		//style.setAlignment(XSSFCellStyle.ALIGN_CENTER);//居中
		if(null==sheet){
			sheet=workbook.createSheet();
		}else{
			workbook.removeSheetAt(sheetNumber);
			sheet=workbook.createSheet();
		}
		int rowNum=0;
		for(String[] value : list){
			XSSFRow row = sheet.getRow(rowNum);
			if(null==row){
				row=sheet.createRow(rowNum);
			}
			for(int i=0;i<value.length;i++){
				if(null!=value[i]){
					String v=value[i];
					XSSFCell cell = row.getCell(i);
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
	 * 写入整个sheet的数据,写入大数据量
	 * @param file 文件名
	 * @param sheetNumber 0开始sheet序号 
	 * @author zhenr
	 * @throws IOException 
	 * */
	public static void writeSheetBig(List<String[]> list,String file,int sheetNumber) throws IOException{		
		FileInputStream fileIn=new FileInputStream(file);
		XSSFWorkbook xworkbook = new XSSFWorkbook(fileIn);
		SXSSFWorkbook workbook = new SXSSFWorkbook(xworkbook,10000);
		SXSSFSheet sheet = workbook.getSheetAt(sheetNumber);
		//XSSFCellStyle style=workbook.createCellStyle();
		//style.setAlignment(XSSFCellStyle.ALIGN_CENTER);//居中
		if(null==sheet){
			sheet=workbook.createSheet();
		}else{
			workbook.removeSheetAt(sheetNumber);
			sheet=workbook.createSheet();
		}
		int rowNum=0;
		for(String[] value : list){
			SXSSFRow row = sheet.getRow(rowNum);
			if(null==row){
				row=sheet.createRow(rowNum);
			}
			for(int i=0;i<value.length;i++){
				if(null!=value[i]){
					String v=value[i];
					SXSSFCell cell = row.getCell(i);
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
	 * 删除某一行，rowNum是行号-1，2007版的shiftRows，startrow参数需要+1
	 * @param file
	 * @param sheetNumber
	 * @param rowNum
	 * @author zhenr
	 * */
	public static void removeRow(String file,int sheetNumber,int rowNum) throws IOException{
		FileInputStream fileIn=new FileInputStream(file);
		XSSFWorkbook workbook = new XSSFWorkbook(fileIn);
		XSSFSheet sheet = workbook.getSheetAt(sheetNumber);
		if(null!=sheet){
			if(null!=sheet.getRow(rowNum))
				sheet.shiftRows(rowNum+1, sheet.getLastRowNum() , -1);
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
		XSSFWorkbook workbook = new XSSFWorkbook(fileIn);
		XSSFSheet sheet = workbook.getSheetAt(sheetNumber);
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
		XSSFWorkbook workbook = new XSSFWorkbook(fileIn);
		XSSFSheet sheet = workbook.getSheetAt(sheetNumber);
		if(null!=sheet){
			XSSFRow row = sheet.getRow(0);
			if(null!=row){
				workbook.close();
				return row.getLastCellNum();
			}else {
				workbook.close();
				return -1;
			}
		}else {
			workbook.close();
			return -1;
		}			
	}
///////////////////////////////////////////////////////////////////	容易异常
	/**
	 * 读取整个sheet的数据 带格式,返回使用queue，这样就带顺序
	 * @param file 文件名
	 * @param sheetNumber 0开始sheet序号 
	 * @author zhenr
	 * @throws IOException 
	 * */
	public static ArrayList<XSSFCell[]> readSheetWithStyle(String file,int sheetNumber) throws IOException{
		ArrayList<XSSFCell[]> list=new ArrayList<XSSFCell[]>();				
		FileInputStream fileIn=new FileInputStream(file);
		XSSFWorkbook workbook = new XSSFWorkbook(fileIn);
		XSSFSheet sheet = workbook.getSheetAt(sheetNumber);			
		if(null!=sheet){
			int rowNum=sheet.getLastRowNum();
			for(int i=0;i<=rowNum;i++){
				XSSFRow row = sheet.getRow(i);
				XSSFCell[] rowResult=null;
				if(null!=row){
					int columnNum=row.getLastCellNum();//列数是实际情况，作为数据index要减1
					//Data.COLUMNNUM=columnNum;
					rowResult=new XSSFCell [columnNum];
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
	 * 写入整个sheet的数据，带格式复制，易超出虚拟机内存,使用queue，这样就带顺序
	 * @param file 文件名
	 * @param sheetNumber 0开始sheet序号 
	 * @author zhenr
	 * @throws IOException 
	 * */
	public static void writeSheetWithStyle(ArrayList<XSSFCell[]> list,String file,int sheetNumber) throws IOException{		
		FileInputStream fileIn=new FileInputStream(file);
		XSSFWorkbook workbook = new XSSFWorkbook(fileIn);
		XSSFSheet sheet = workbook.getSheetAt(sheetNumber);
		//XSSFCellStyle style=workbook.createCellStyle();
		//style.setAlignment(XSSFCellStyle.ALIGN_CENTER);//居中
		if(null==sheet){
			sheet=workbook.createSheet();
		}
		int rowNum=0;
		for(XSSFCell[] value : list){
			if(value!=null) {
				XSSFRow row = sheet.getRow(rowNum);
				if(null==row){
					row=sheet.createRow(rowNum);
				}
				for(int i=0;i<value.length;i++){
					if(null!=value[i]){
						XSSFCell cell = row.getCell(i);
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
        }				
		FileOutputStream fileOut = new FileOutputStream(file);
		workbook.write(fileOut);
		workbook.close();
		fileIn.close();
		fileOut.close();	
	}
	
	/**
	 * 写入数据到一个sheet，从指定起始行、起始列开始写，并使用起始行格式写入
	 * @param list 需写入的数据，使用Queue类型以保持顺序；file目标文件，sheetNumber 指定sheet号，rowNum 起始行号 0开始，colmunNum 起始列号 0开始
	 * @throws IOException 
	 * */
	public static void writeSheet(ArrayList<Object[]> list,String file,int sheetNumber,int rowNum,int colmunNum) throws IOException {
		FileInputStream fileIn=new FileInputStream(file);
		XSSFWorkbook workbook = new XSSFWorkbook(fileIn);
		XSSFSheet sheet = workbook.getSheetAt(sheetNumber);
		if(null==sheet){
			fileIn.close();
			workbook.close();
			return;
		}
		/*先记录起始行格式
		XSSFRow startRow = sheet.getRow(rowNum);
		XSSFCell[] styleRow=null;//每格都放入一个数组
		if(null!=startRow){
			int columnNum=startRow.getLastCellNum();//列数是实际情况，作为数据index要减1
			styleRow=new XSSFCell [columnNum];
			for(int j=0;j<columnNum;j++){
				styleRow[j]=startRow.getCell(j);
			}
		}
		*/
		for(Object[] value : list){
			if(value!=null) {
				XSSFRow row = sheet.getRow(rowNum);
				if(null==row){
					row=sheet.createRow(rowNum);
				}
				int column=colmunNum;
				for(int i=0;i<value.length;i++){
					if(null!=value[i]){
						XSSFCell cell = row.getCell(column);
						if(null==cell){
							cell=row.createCell(i);
						}
						if (value[i] instanceof Integer) {
							cell.setCellValue(((Integer) value[i]).intValue());
						}else if (value[i] instanceof Double) {
							cell.setCellValue(((Double) value[i]).doubleValue());
						}else if (value[i] instanceof String) {
							cell.setCellValue((String) value[i]);
						}
						//cell.setCellStyle(styleRow[column].getCellStyle());
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
	 * 设置startRow开始的行的Style为sampleRow的Style,startRowNo开始写的列
	 * @throws IOException 
	 * */
	public static void setStyleFromRow(String file,int sheetNumber,int sampleRowNo,int startRowNo) throws IOException {
		FileInputStream fileIn=new FileInputStream(file);
		XSSFWorkbook workbook = new XSSFWorkbook(fileIn);
		XSSFSheet sheet = workbook.getSheetAt(sheetNumber);
		if(null==sheet){
			fileIn.close();
			workbook.close();
			return;
		}
		//先记录起始行格式
		XSSFRow sampleRow = sheet.getRow(sampleRowNo);
		XSSFCell[] styleRow=null;//每格都放入一个数组
		if(null!=sampleRow){
			int columnNum=sampleRow.getLastCellNum();//列数是实际情况，作为数据index要减1
			styleRow=new XSSFCell [columnNum];
			for(int j=0;j<columnNum;j++){
				styleRow[j]=sampleRow.getCell(j);
			}
		}else {
			fileIn.close();
			workbook.close();
			return;
		}
		if(null!=styleRow&&styleRow.length>0) {
			int rowNum=sheet.getLastRowNum();
			for(int i=startRowNo;i<=rowNum;i++){
				XSSFRow row = sheet.getRow(i);
				if(null!=row){
					int columnNum=row.getLastCellNum();//列数是实际情况，作为数据index要减1					
					for(int j=0;j<columnNum;j++){
						XSSFCell cell = row.getCell(j);
						if(null==cell){
							cell=row.createCell(j);
						}
						if(null!=styleRow[j]&&null!=styleRow[j].getCellStyle()) {
							cell.setCellStyle(styleRow[j].getCellStyle());
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
	public static String readCell(String file,int sheetNumber,int sampleRowNo,int startRowNo) throws IOException {
		String result="";
		FileInputStream fileIn=new FileInputStream(file);
		XSSFWorkbook workbook = new XSSFWorkbook(fileIn);
		XSSFSheet sheet = workbook.getSheetAt(sheetNumber);
		if(null==sheet){
			fileIn.close();
			workbook.close();
			return null;
		}
		XSSFRow row = sheet.getRow(sampleRowNo);
		if(null==row){
			fileIn.close();
			workbook.close();
			return null;
		}
		XSSFCell cell=row.getCell(startRowNo);
		if(null==cell){
			fileIn.close();
			workbook.close();
			return null;
		}
		
		if(cell.getCellType()==0){
			DecimalFormat decimalFormat = new DecimalFormat("##0");//double格式化设置,不要科学计数法
			result=decimalFormat.format(cell.getNumericCellValue());
		}
		else if(cell.getCellType()==1){
			result=cell.getStringCellValue();
		}
		fileIn.close();
		workbook.close();
		return result;
	}
	
	/**
	 * 带格式修改一格
	 * @throws IOException 
	 * */
	public static void editCell(Object value,String file,int sheetNumber,int sampleRowNo,int startRowNo) throws IOException {
		FileInputStream fileIn=new FileInputStream(file);
		XSSFWorkbook workbook = new XSSFWorkbook(fileIn);
		XSSFSheet sheet = workbook.getSheetAt(sheetNumber);
		if(null==sheet){
			fileIn.close();
			workbook.close();
			return;
		}
		XSSFRow row = sheet.getRow(sampleRowNo);
		if(null==row){
			fileIn.close();
			workbook.close();
			return;
		}
		XSSFCell cell=row.getCell(startRowNo);
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
