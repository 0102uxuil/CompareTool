package com.lx.tools;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
//import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

public class ExcelReader {
	
	private ArrayList<ExcelRow> al;
	
	private InputStream is;
    private POIFSFileSystem fs;
    private HSSFWorkbook wb;
    private HSSFSheet sheet;
	
	public ExcelReader(String filepath){
		
		al = new ArrayList<ExcelRow>();
		
		try {
			is = new FileInputStream(filepath);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
        try {
			fs = new POIFSFileSystem(is);
			wb = new HSSFWorkbook(fs);
			sheet = wb.getSheetAt(0);
			int RowNum = sheet.getLastRowNum()+1;
//			int ColNum = sheet.getRow(0).getPhysicalNumberOfCells();
			ExcelRow er;
			for(int i=0; i<RowNum; i++){
				System.out.println("s：" + i);
				er = new ExcelRow();
				er.setRiqi(sheet.getRow(i).getCell(0).getDateCellValue());
				er.setChepaihao(sheet.getRow(i).getCell(1).getStringCellValue());
				er.setShengshu(sheet.getRow(i).getCell(2).getNumericCellValue());
				er.setYoujia(sheet.getRow(i).getCell(3).getNumericCellValue());
				er.setShifoupeidui(0);//没有匹配的，为0
				al.add(er);
			}
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	public ArrayList<ExcelRow> getAl() {
		return al;
	}

	public void setAl(ArrayList<ExcelRow> al) {
		this.al = al;
	}
	
}
