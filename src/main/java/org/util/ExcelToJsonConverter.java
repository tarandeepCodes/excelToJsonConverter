package org.util;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.fasterxml.jackson.databind.ObjectMapper;

import java.io.*;
import java.util.*;

public class ExcelToJsonConverter {

    public static void main(String[] args) throws Exception {
        if (args.length != 2) {
            System.out.println("Usage: java -jar excel-to-json.jar <input.xlsx> <output.json>");
            return;
        }

        File excelFile = new File(args[0]);
        File jsonFile = new File(args[1]);

        List<Map<String, String>> jsonData = convertExcelToJson(excelFile);
        ObjectMapper mapper = new ObjectMapper();
        mapper.writerWithDefaultPrettyPrinter().writeValue(jsonFile, jsonData);

        System.out.println("Successfully converted Excel to JSON!");
    }

    public static List<Map<String, String>> convertExcelToJson(File file) throws IOException {
        List<Map<String, String>> data = new ArrayList<>();

        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            Row headerRow = sheet.getRow(0);
            int columnCount = headerRow.getPhysicalNumberOfCells();

            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                Map<String, String> rowData = new LinkedHashMap<>();
                for (int j = 0; j < columnCount; j++) {
                    String header = headerRow.getCell(j).getStringCellValue();
                    Cell cell = row.getCell(j);
                    String value = (cell == null) ? "" : cell.toString();
                    rowData.put(header, value);
                }
                data.add(rowData);
            }
        }

        return data;
    }
}