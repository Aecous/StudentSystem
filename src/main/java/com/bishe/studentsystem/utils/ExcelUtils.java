package com.bishe.studentsystem.utils;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

public class ExcelUtils {

    public static void generateExcel(ArrayList<String> headers, ArrayList<ArrayList<String>> data, String filePath) throws IOException {
        // 创建一个新的Excel工作簿（对应一个.xlsx文件）
        Workbook workbook = new XSSFWorkbook();

        // 创建一个工作表
        Sheet sheet = workbook.createSheet("Sheet1");

        // 创建表头行
        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < headers.size(); i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers.get(i));
        }

        // 填充数据行
        for (int rowIndex = 0; rowIndex < data.size(); rowIndex++) {
            ArrayList<String> rowData = data.get(rowIndex);
            Row dataRow = sheet.createRow(rowIndex + 1);
            for (int colIndex = 0; colIndex < rowData.size(); colIndex++) {
                Cell cell = dataRow.createCell(colIndex);
                cell.setCellValue(rowData.get(colIndex));
            }
        }

        // 将工作簿写入到文件
        try (FileOutputStream outputStream = new FileOutputStream(filePath)) {
            workbook.write(outputStream);
        }
    }


}

