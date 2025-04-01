package org.example;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;

//TIP To <b>Run</b> code, press <shortcut actionId="Run"/> or
// click the <icon src="AllIcons.Actions.Execute"/> icon in the gutter.
public class Main {

    public static void main(String[] args) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(new File("src/main/resources/spreadsheet.xlsx")));

        XSSFSheet sheet = workbook.getSheetAt(0);
        CreationHelper createHelper = workbook.getCreationHelper();

        for(int j = 1; j < 4; j ++){
            for(int i = 1; i < 14; i++) {
                Row initRow = sheet.getRow(i);
                Row row = sheet.createRow(5 + 3*i);

                System.out.println(initRow.getCell(j).getStringCellValue());
                switch (initRow.getCell(j).getCellType()) {
                    case STRING:
                        row.createCell(j).setCellValue(createHelper.createRichTextString(initRow.getCell(j).getStringCellValue()));
                        break;

                    case NUMERIC:
                        row.createCell(j).setCellValue(initRow.getCell(j).getNumericCellValue());
                        break;
                }
            }
        }


    }

}