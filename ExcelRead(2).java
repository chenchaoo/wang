package one;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExcelRead {

	private static final Map<String, Map<String, String>> confPoListMap = ReadConf.getConfMap();

	@SuppressWarnings("unchecked")
	public static boolean generateWorkbook(String path, Map<String, List<Object>> mapList) {
		Workbook workbook = new XSSFWorkbook();
		for (String string : mapList.keySet()) {
        	List<Object> objList = mapList.get(string);
        	for (Object object : objList) {
        		if (object instanceof HashMap) {
        			Map<String, List<Object>> mapChile = (Map<String, List<Object>>) object;
        			String sheetName = getSheetName(mapChile);
        			createSheet(path, sheetName, mapChile, workbook);
        		}
			}
        	
		}
		boolean isCorrect = false;
		File file = new File(path);
		if (file.exists()) {
			file.deleteOnExit();
			file = new File(path);
		}
		OutputStream outputStream = null;
		try {
			outputStream = new FileOutputStream(file);
			workbook.write(outputStream);
			isCorrect = true;
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			try {
				if (null != outputStream) {
					outputStream.close();
				}
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		return isCorrect;
	}

	private static String getSheetName(Map<String, List<Object>> mapChile) {
		List<String> list = new ArrayList<>(mapChile.keySet());
		String sheetName = "";
		if(!list.isEmpty()){
			String sheetKey = list.get(0);
			sheetName = sheetKey.substring(sheetKey.indexOf("=")+1);
		}
		return sheetName;
	}

	private static void createSheet(String path, String name, Map<String, List<Object>> mapList, Workbook workbook) {
		Sheet sheet = workbook.createSheet(name);
		sheet.setDefaultColumnWidth((short) 25);
		Map<String, CellStyle> styles = createStyles(workbook);
		int rowNum = 0;
		int cellNum = 0;
		read(mapList, rowNum, cellNum, sheet, styles);
	}

	private static Map<String, CellStyle> createStyles(Workbook wb) {
		Map<String, CellStyle> styles = new HashMap<>();
		CellStyle cellStyle = wb.createCellStyle();
		cellStyle.setAlignment(HorizontalAlignment.CENTER);
		cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		cellStyle.setWrapText(true);
		cellStyle.setBorderRight(BorderStyle.THIN);
		cellStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
		cellStyle.setBorderLeft(BorderStyle.THIN);
		cellStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		cellStyle.setBorderTop(BorderStyle.THIN);
		cellStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
		cellStyle.setBorderBottom(BorderStyle.THIN);
		cellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		styles.put("cell", cellStyle);

		return styles;
	}

	private static String getKeyConvert(String key) {
		String newKey = key.substring(0, key.indexOf("="));
		Map<String, String> map = confPoListMap.get(newKey);
		if (map != null && !map.isEmpty()) {
			String string = map.get(newKey);
			return string == null ? newKey : string;
		}
		return newKey;
	}

	private static String getValueConvert(String value) {
		String newKey = value.substring(0, value.indexOf("="));
		String newValue = value.substring(value.indexOf("=") + 1);
		Map<String, String> map = confPoListMap.get(newKey);
		if (map != null && !map.isEmpty()) {
			String string = map.get(newValue);
			String oldChar = "";
			if (string == null) {
				for (String key : map.keySet()) {
					if (newValue.contains(key)) {
						oldChar = key;
						string = map.get(key);
					}
				}
				return string == null ? newValue : newValue.replace(oldChar, string);
			} else {
				return string;
			}
		}
		return newValue;
	}

	public static int read(Map<String, List<Object>> mapList, int rowNum, int cellNum, Sheet sheet,
			Map<String, CellStyle> styles) {
		for (String key : mapList.keySet()) {
			Row row = sheet.createRow(rowNum);
			Cell cell = row.createCell(cellNum);
			cell.setCellStyle(styles.get("cell"));
			cell.setCellValue(getKeyConvert(key));
			cellNum += 1;
			Cell valueCell = row.createCell(cellNum);
			valueCell.setCellStyle(styles.get("cell"));
			valueCell.setCellValue(getValueConvert(key));
			List<Object> list = mapList.get(key);
			cellNum += 1;
			for (int i = 0; i < list.size(); i++) {
				Object object = list.get(i);
				if (object instanceof HashMap) {
					@SuppressWarnings("unchecked")
					Map<String, List<Object>> mapChile = (Map<String, List<Object>>) object;
					rowNum = read(mapChile, rowNum + 1, cellNum, sheet, styles);
				} else {
					String newKey = object.toString().substring(0, object.toString().indexOf("="));
					if("el".equals(newKey)){
						continue;
					}
					rowNum += 1;
					Row rowChild = sheet.createRow(rowNum);
					Cell cellChild = rowChild.createCell(cellNum);
					cellChild.setCellStyle(styles.get("cell"));
					cellChild.setCellValue(getKeyConvert(object.toString()));
					Cell cellValueChild = rowChild.createCell(cellNum+1);
					cellValueChild.setCellStyle(styles.get("cell"));
					cellValueChild.setCellValue(getValueConvert(object.toString()));
				}
			}
		}
		return rowNum;
	}
}
