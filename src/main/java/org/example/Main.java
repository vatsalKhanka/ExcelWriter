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
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;

//TIP To <b>Run</b> code, press <shortcut actionId="Run"/> or
// click the <icon src="AllIcons.Actions.Execute"/> icon in the gutter.
public class Main {

    public static void main(String[] args) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(new File("src/main/resources/spreadsheet.xlsx")));

        XSSFSheet sheet = workbook.getSheetAt(0);
        CreationHelper createHelper = workbook.getCreationHelper();

        for(int j = 2; j < 4; j ++){
            for(int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
                Row initRow = sheet.getRow(i);

                switch (initRow.getCell(j).getCellType()) {
                    case STRING:
                        System.out.println(createHelper.createRichTextString(initRow.getCell(j).getStringCellValue()));
                        Cell cell0 = initRow.createCell(j*3 + 3);
                        Cell cell1 = initRow.createCell(j*3 + 4);
                        cell0.setCellValue(initRow.getCell(1).getStringCellValue());
                        cell1.setCellValue(createHelper.createRichTextString(initRow.getCell(j).getStringCellValue()));
                        break;

                    case NUMERIC:
                        System.out.println(initRow.getCell(j).getNumericCellValue());
                        Cell cell2 = initRow.createCell(j*3 + 3);
                        Cell cell3 = initRow.createCell(j*3 + 4);
                        cell2.setCellValue(initRow.getCell(1).getStringCellValue());
                        cell3.setCellValue(initRow.getCell(j).getNumericCellValue());
                        break;
                }
            }
        }

        FileOutputStream out = new FileOutputStream("src/main/resources/newspreadsheet.xlsx");
        workbook.write(out);
        out.close();

    }

}