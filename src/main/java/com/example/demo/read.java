package com.example.demo;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Properties;
import java.util.stream.Stream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.monitorjbl.xlsx.StreamingReader;

import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * 遞迴讀取某個目錄下的所有檔案
 * 
 * @author 超越
 * @Date 2016年12月5日,下午4:04:59
 * @motto 人在一起叫聚會，心在一起叫團隊
 * @Version 1.0
 */

public class read {
	private final static Logger logger = LoggerFactory.getLogger(read.class);
	static List<String> queryColArray;// 要抓取的欄位
	static String excelFolderPath; // excel資料夾位置
	static String systemName; // excel檔名
	static String sheetName; // 分頁名
	static int col; // columns 位置
	// (CellIndex,HeaderName)
	static Map<Integer, String> HeaderName = new HashMap<Integer, String>();

	public static void main(String[] args) {
		test("D:/si1204/Desktop/writeY/all_copybook_20210820");
	}

	private static void test(String fileDir) {

		List<File> fileList = new ArrayList<File>();
		File file = new File(fileDir);
		File[] files = file.listFiles();// 獲取目錄下的所有檔案或資料夾
		if (files == null) {// 如果目錄為空，直接退出
			return;
		}
		// 遍歷，目錄下的所有檔案
		for (File f : files) {
			if (f.isFile()) {
				fileList.add(f);
			} else if (f.isDirectory()) {
//	System.out.println(f.getAbsolutePath()); 
				test(f.getAbsolutePath());
			}
		}
		for (File f1 : fileList) {

			String copybookpath = fileDir + "\\" + f1.getName();
			String copybookname = f1.getName();
			readcopybook(copybookpath, copybookname);
		}

	}

	public static void readcopybook(String copybookpath, String copybookname) {
		FileReader reader;
		String str1 = null;
		try {
			reader = new FileReader(copybookpath);
			BufferedReader br = new BufferedReader(reader);
//System.out.println(copybookpath);
			
			//comp3的條件
			while ((str1 = br.readLine()) != null) {
				if (str1.contains("S9") && str1.contains("COMP")) {
//					System.out.println(copybookpath + "11111111111111111111");
					try {

						readexcel(copybookname);
					} catch (Exception e) {
						e.printStackTrace();
					}
					break;
				} else if (str1.contains("PIC 9") && str1.contains("COMP")) {
//					System.out.println(copybookname + "22222222222222222222");
					try {
						readexcel(copybookname);
					} catch (Exception e) {
						e.printStackTrace();
					}
					break;
				} else if (str1.contains("PIC  9") && str1.contains("COMP")) {
//					System.out.println(copybookname + "33333333333333333333");
					try {
						readexcel(copybookname);
					} catch (Exception e) {
						e.printStackTrace();
					}
					break;
				}

			}

			br.close();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	public static void readexcel(String copybookname) throws Exception {
		Properties pro = new Properties();
		String config = "config.properties";

		pro.load(new FileInputStream(config));
		excelFolderPath = pro.getProperty("excelDir");
		queryColArray = Arrays.asList(pro.getProperty("queryColArray").split(","));

//		System.out.println(excelFolderPath);
		File excelFolder = new File(excelFolderPath);
		List<Map<String, String>> excelInfoList = null;

		for (File file : excelFolder.listFiles()) {

			Workbook wb = getExcelFile(file.getPath());
			if (wb == null) {
				throw new Exception("讀取失敗");
			}
			// 解析Excel to List
			excelInfoList = parseExcel(wb);
//			System.out.println(excelInfoList);
			// excel檔名
			
			systemName = file.getName();
			wb.close();

		}
		writeexcel(excelInfoList, copybookname, excelFolderPath);

	}

	public static void writeexcel(List<Map<String, String>> excelInfoList, String copybookname,
			String excelFolderPath) {
//		System.out.println(copybookname);
		
		
		for (Map<String, String> text : excelInfoList) {
			//路徑切割，比對
			String lastItem = Stream.of(text.get("Copybook").split("/")).reduce((first, last) -> last).get();
			if (lastItem.equals(copybookname)) {
				System.out.println(lastItem);
				System.out.println(text.get("rowindex"));
				int rownum = Integer.valueOf(text.get("rowindex"));

				Workbook workbook;
				try {
					//寫入EXCEL
					workbook = new XSSFWorkbook(excelFolderPath + "\\" + systemName);
					Sheet sheet1 = workbook.getSheet(sheetName);
					Cell writeCell = sheet1.getRow(rownum).getCell(col);
					
					if (writeCell == null) {
						writeCell = sheet1.getRow(rownum).createCell(col);
					}
					writeCell.setCellValue("Y");
					FileOutputStream fos = new FileOutputStream(excelFolderPath + "\\" + systemName, true);
					workbook.write(fos);
					workbook.close();
					fos.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}

			}

		}
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
	public static Workbook getExcelFile(String path) {
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

	public static List<Map<String, String>> parseExcel(Workbook workbook) {
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
							//給colums位置
							if (cell.getStringCellValue().equals("Isbinary")) {
								col = cell.getColumnIndex();
							}
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

	public static Map<String, String> convertRowToData(Row row, int dnsIndex) {
		Map<String, String> excelDateMap = new HashMap<String, String>();

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

				excelDateMap.put("rowindex", String.valueOf(cell.getRowIndex()));
			}
		}

		return excelDateMap;
	}

}
