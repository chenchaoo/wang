package one;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
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
    public static boolean generateWorkbook(String path, String name, Map<String,List<Object>> mapList) {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet(name);
        sheet.setDefaultColumnWidth((short) 15);
        Map<String, CellStyle> styles = createStyles(workbook);
        int rowNum = 0;
        int cellNum = 0;
        read(mapList, rowNum, cellNum, sheet, styles);
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
    
    public static int read(Map<String,List<Object>> mapList,int rowNum,int cellNum,Sheet sheet,Map<String, CellStyle> styles){
        for (String key : mapList.keySet()) {
            Row row = sheet.createRow(rowNum);
            Cell cell = row.createCell(cellNum);
            cell.setCellStyle(styles.get("cell"));
            cell.setCellValue(key);
            List<Object> list = mapList.get(key);
            rowNum -= 1;
            cellNum += 1;
            boolean falg = false;
            for (int i = 0; i < list.size(); i++) {
                Object object = list.get(i);
                if(object instanceof HashMap){
                   @SuppressWarnings("unchecked")
                   Map<String,List<Object>> mapChile = (Map<String, List<Object>>) object;
                   rowNum = read(mapChile, rowNum+1, cellNum, sheet, styles);
                }else{
                    rowNum += 1;
                    if(!falg){
                        Cell cellChild = row.createCell(cellNum);
                        cellChild.setCellStyle(styles.get("cell"));
                        cellChild.setCellValue(object.toString());
                        falg = true;
                    }else{
                        Row rowChild = sheet.createRow(rowNum);
                        Cell cellChild = rowChild.createCell(cellNum);
                        cellChild.setCellStyle(styles.get("cell"));
                        cellChild.setCellValue(object.toString());
                    }
                }
            }
        }
        return rowNum;
    }
}
