package com.example.demo;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Properties;
import java.util.Scanner;
import java.util.stream.Collector;
import java.util.stream.Collectors;

import org.apache.commons.io.filefilter.AbstractFileFilter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.monitorjbl.xlsx.StreamingReader;


public class readexcel2 {

	private final static Logger logger = LoggerFactory.getLogger(readexcel2.class);
	List<String> queryColArray;// 要抓取的欄位
	String excelFolderPath; // excel資料夾位置
	String destFileFolderPath; // output位置
	String systemName;
	String sheetName;

	// (CellIndex,HeaderName)
	Map<Integer, String> HeaderName = new HashMap<Integer, String>();

	readexcel2() {
		Properties pro = new Properties();
		String config = "config.properties";
		logger.info(config);
		try {
			pro.load(new FileInputStream(config));
			excelFolderPath = pro.getProperty("excelDir");
//			destFileFolderPath = pro.getProperty("destFile");
//			logger.info(" Excel資料夾位置:{}\n  輸出資料夾位置:{}", excelFolderPath, destFileFolderPath);

			queryColArray = Arrays.asList(pro.getProperty("queryColArray").split(","));
			logger.info("需要抓取的欄位 " + queryColArray);

		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

	public void readExcelStart() throws Exception {
		// 取資料夾
		File excelFolder = new File(excelFolderPath);
//		logger.info("excelDir:{} 有 {} 個Excel檔案", excelFolderPath, excelFolder.list().length);
		List<Map<String, String>> excelInfoList = null;
		for (File file : excelFolder.listFiles()) {
			logger.info("開始讀取 " + file.getName());

			Workbook wb = getExcelFile(file.getPath());
			if (wb == null) {
				logger.info(file.getPath() + "讀取失敗");
				throw new Exception("讀取失敗");
			}
			logger.info(file.getName() + " 讀取完成");
			// 解析Excel to List
			logger.info("開始解析 " + file.getName());
			excelInfoList = parseExcel(wb);
			logger.info("解析 " + file.getName() + " 完成");
			// excel檔名
			// EX: excel檔名 帳務作業流程清單(BANK)_1090430 取 帳務作業流程清單(BANK)
			systemName = file.getName();
			logger.info("檔名 " + sheetName);
//			System.out.println(excelInfoList);

			logger.info(systemName + " 完成");

		}
//		findPreviousAndNext(excelInfoList);
//			System.out.println(excelInfoList);
		readcopybook(excelInfoList);
	}

	public void readcopybook(List<Map<String, String>> excelInfoList) {
//		for (Map<String, String> text : excelInfoList) {
//			System.out.println(text);
//		}
		
		
			List<Map<String, String>> filterList;
			filterList = excelInfoList.stream()
					.filter(afFilter -> afFilter.get("Conversion mode").equals("COPYBOOK based")&&
					afFilter.get("Sprint").equals("sprint1")&&
					afFilter.get("In Scope").equals("Yes"))
					.collect(Collectors.toList());
			System.out.println(filterList.isEmpty());
			for(Map<String, String>text:filterList) {
			System.out.println(text);
			}
	
	
	}
	
	
	
	
	public void findPreviousAndNext(List<Map<String, String>> excelInfoList) {
		Scanner scanner = new Scanner(System.in);
		System.out.println("輸入job name");
		String keyjob = scanner.nextLine();
		logger.info("要查詢的job name " + keyjob);

		System.out.println("============往前找=======================");
		for (Map<String, String> text : excelInfoList) {
//				System.out.println(text);
			if (text.get("JOB").equals(keyjob)) {
				System.out.println("現在的AD NAME: " + text.get("AD") + '\t' + "PreviousJOB: " + text.get("PreviousJOB")
						+ '\t' + "NextJOB: " + text.get("NextJOB"));
				break;
			} else {
				System.out.println("前一筆的AD NAME: " + text.get("AD") + '\t' + "PreviousJOB: " + text.get("PreviousJOB")
						+ '\t' + "NextJOB: " + text.get("NextJOB"));

			}
		}

		System.out.println("============往後找=======================");
		for (int i = 0; i < excelInfoList.size(); i++) {
			if (excelInfoList.get(i).get("JOB").equals(keyjob)) {
				System.out.println("現在的AD NAME: " + excelInfoList.get(i).get("AD") + '\t' + "PreviousJOB: "
						+ excelInfoList.get(i).get("PreviousJOB") + '\t' + "NextJOB: "
						+ excelInfoList.get(i).get("NextJOB"));

				for (int j = i + 1; j < excelInfoList.size(); j++) {
					System.out.println("下一個AD NAME: " + excelInfoList.get(j).get("AD") + '\t' + "PreviousJOB: "
							+ excelInfoList.get(j).get("PreviousJOB") + '\t' + "NextJOB: "
							+ excelInfoList.get(j).get("NextJOB"));

				}
			}

		}

		scanner.close();

	}

	/**
	 * 讀取excel檔案
	 * 
	 * Workbook(Excel本體)、Sheet(內部頁面)、Row(頁面之行(橫的))、Cell(行內的元素)
	 * 
	 * 
	 * @param path excel檔案路徑
	 * @return excel內容
	 */
	public Workbook getExcelFile(String path) {
		Workbook wb = null;
		try {
			if (path == null) {
				return null;
			}
			String extString = path.substring(path.lastIndexOf(".")).toLowerCase();
			FileInputStream in = new FileInputStream(path);
			wb = StreamingReader.builder().rowCacheSize(100)// 存到記憶體行數，預設10行。
					.bufferSize(2048)// 讀取到記憶體的上限，預設1024
					.open(in);
		} catch (FileNotFoundException e) {
			logger.info(e.toString());
			e.printStackTrace();
		}

		return wb;
	}

	/**
	 * 解析Sheet
	 * 
	 * @param workbook Excel檔案
	 * @return 整個Sheet的資料
	 */
	public List<Map<String, String>> parseExcel(Workbook workbook) {
		// Sheet的資料
		List<Map<String, String>> excelDataList = new ArrayList<>();
		// 存放DNS欄位的欄位號
		int dnsIndex = 0;
		// 遍歷每一個sheet
		for (int sheetNum = 0; sheetNum < workbook.getNumberOfSheets(); sheetNum++) {
			Sheet sheet = workbook.getSheetAt(sheetNum);
			boolean rowNum = true;

			sheetName = sheet.getSheetName();
			// 開始讀取sheet
			for (Row row : sheet) {
				// 先取header
				if (rowNum) {
					for (Cell cell : row) {
						if (queryColArray.contains(cell.getStringCellValue())) {
							HeaderName.put(cell.getColumnIndex(), cell.getStringCellValue());
						}
					}
					rowNum = false;
					continue;
				}

				// 解析Row的資料
				excelDataList.add(convertRowToData(row, dnsIndex));
			}
		}
		return excelDataList;
	}

	/**
	 * 解析ROW
	 * 
	 * @param row      資料行
	 * @param firstRow 標頭
	 * @param dnsIndex Dns的列數
	 * @return 整row的欄位
	 */
	public Map<String, String> convertRowToData(Row row, int dnsIndex) {
		Map<String, String> excelDateMap = new HashMap<String, String>();

//		Scanner scanner = new Scanner(System.in);
//		System.out.println("輸入job name");
//		String keyjob = scanner.nextLine();

		for (Object key : HeaderName.keySet()) {
			for (Cell cell : row) {

				// 1.
				int headerNameIndex = (int) key;

				cell = row.getCell(headerNameIndex, MissingCellPolicy.CREATE_NULL_AS_BLANK);
				if (cell == null) {
					cell.setCellValue("Empty");
				}
				// 2.再去抓Header的欄位名稱
				String firstRowName = HeaderName.get(key);
				excelDateMap.put(firstRowName, cell.getStringCellValue());

				excelDateMap.put(firstRowName, cell.getStringCellValue());
			}
		}

		return excelDateMap;
	}

}
