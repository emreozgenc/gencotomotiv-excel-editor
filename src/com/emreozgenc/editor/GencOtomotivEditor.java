package com.emreozgenc.editor;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class GencOtomotivEditor {

    private FileInputStream inputStream;
    private FileOutputStream outputStream;

    private Workbook wb;
    private Workbook nwb;

    public GencOtomotivEditor(String word, String rWord, String path, String sheetName, int param, int column) throws FileNotFoundException, IOException {
        inputStream = new FileInputStream(path);
        outputStream = new FileOutputStream("genc.xlsx");

        if (path.endsWith("xls")) {
            wb = new HSSFWorkbook(inputStream);
        } else if (path.endsWith("xlsx")) {
            wb = new XSSFWorkbook(inputStream);
        }

        nwb = new XSSFWorkbook();

        Sheet sheet = wb.getSheet(sheetName);
        Sheet nsheet = nwb.createSheet(sheetName);

        for (int i = 0; i < sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            Row nrow = nsheet.createRow(i);

            Cell cell = row.getCell(column - 1);
            Cell ncell = nrow.createCell(column - 1);
            
            String stockString = cell.getStringCellValue();
            
            switch(param) {
                case 0:
                    stockString = stockString.replace(word, rWord);
                    ncell.setCellValue(stockString);
                    break;
                case 1:
                    stockString = rWord + stockString;
                    ncell.setCellValue(stockString);
                    break;
                case 2:
                    stockString = stockString + rWord;
                    ncell.setCellValue(stockString);
                    break;          
            }
        }
        
        wb.close();
        inputStream.close();
        
        nwb.write(outputStream);
        nwb.close();
        outputStream.close();
        
        EditorLauncher.launcher.sendSuccessMessage();
    }
}
