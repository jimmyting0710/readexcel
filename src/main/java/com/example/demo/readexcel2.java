package com.example.demo;

import java.io.FileInputStream;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

public class readexcel2 {
	public void readExcel() {
		FileInputStream fis;
		POIFSFileSystem fs;
		HSSFWorkbook wb;
		String filePath = "D:/si1204/Desktop/ReadExcel/read";
		try {

			fis = new FileInputStream(filePath);
			fs = new POIFSFileSystem(fis);
			wb = new HSSFWorkbook(fs);
			HSSFSheet sheet = wb.getSheetAt(0);
// 取得Excel第一個sheet(從0開始) 
			HSSFCell cell;

			// getPhysicalNumberOfRows這個比較好
			for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
// 由於第 0 Row 為 title, 故 i 從 1 開始 
				HSSFRow row = sheet.getRow(i); // 取得第 i Row
				if (row != null) {
					int j = 0;
					for (; j < 9; j++) { // 看資料需要的欄數
						cell = row.getCell(j);
						System.out.println(cell.toString());// 取出j列j行的值
					}
				}

			}
			fis.close();// 懶的寫到外面去了...
		} catch (java.io.IOException e) {
			e.printStackTrace();
		}
	}

}
