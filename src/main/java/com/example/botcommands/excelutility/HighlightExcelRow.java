package com.example.botcommands.excelutility;

import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.BooleanValue;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

@BotCommand
@CommandPkg(
        name = "HighlightExcelRow",
        label = "Highlight Excel Row",
        node_label = "Highlight Excel rows in {{filePath}}",
        description = "Colors specified rows in an Excel file with the given HEX color",
        icon = "icons/Microsoft_Office_Excel_Logo.png",
        return_type = DataType.BOOLEAN,
        return_label = "Success",
        return_description = "Returns true if successful"
)
public class HighlightExcelRow {

    @Execute
    public Value<Boolean> action(
            @Idx(index = "1", type = AttributeType.FILE)
            @Pkg(label = "Excel File Path")
            String filePath,

            @Idx(index = "2", type = AttributeType.TEXT)
            @Pkg(label = "Row Numbers (comma-separated, e.g., 2,5,7)")
            String rowNumbersStr,

            @Idx(index = "3", type = AttributeType.TEXT)
            @Pkg(label = "Start Column Number (e.g., 1)")
            String startColumnStr,

            @Idx(index = "4", type = AttributeType.TEXT)
            @Pkg(label = "Column Count (e.g., 5)")
            String columnCountStr,

            @Idx(index = "5", type = AttributeType.TEXT)
            @Pkg(label = "Color (HEX format, e.g., #90EE90)")
            String colorHex,

            @Idx(index = "6", type = AttributeType.TEXT)
            @Pkg(label = "Sheet Name (optional)")
            String sheetName
    ) throws Exception {

        if (filePath == null || filePath.isEmpty()) {
            throw new Exception("File path is required.");
        }

        // 🧮 Safely parse numeric values
        int startCol = 1;
        int totalCols = 1;
        try {
            if (startColumnStr != null && !startColumnStr.trim().isEmpty()) {
                startCol = (int) Double.parseDouble(startColumnStr.trim());
            }
            if (columnCountStr != null && !columnCountStr.trim().isEmpty()) {
                totalCols = (int) Double.parseDouble(columnCountStr.trim());
            }
        } catch (NumberFormatException e) {
            throw new Exception("Invalid number format for Start Column or Column Count.");
        }

        int endCol = startCol + totalCols - 1;

        try (InputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = (sheetName != null && !sheetName.isEmpty())
                    ? workbook.getSheet(sheetName)
                    : workbook.getSheetAt(0);

            if (sheet == null && sheetName != null) {
                throw new Exception("Sheet not found: " + sheetName);
            }

            // Parse HEX color
            if (colorHex.startsWith("#")) {
                colorHex = colorHex.substring(1);
            }
            int r = Integer.parseInt(colorHex.substring(0, 2), 16);
            int g = Integer.parseInt(colorHex.substring(2, 4), 16);
            int b = Integer.parseInt(colorHex.substring(4, 6), 16);

            XSSFCellStyle style = (XSSFCellStyle) workbook.createCellStyle();
            XSSFColor color = new XSSFColor(new java.awt.Color(r, g, b), null);
            style.setFillForegroundColor(color);
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            // Handle row numbers safely
            if (rowNumbersStr == null || rowNumbersStr.trim().isEmpty()) {
                throw new Exception("Row numbers are required.");
            }

            String[] rowTokens = rowNumbersStr.split(",");
            for (String token : rowTokens) {
                token = token.trim();
                if (token.isEmpty()) continue;

                int rowNum;
                try {
                    rowNum = (int) Double.parseDouble(token);
                } catch (NumberFormatException e) {
                    throw new Exception("Invalid row number: " + token);
                }

                Row row = sheet.getRow(rowNum - 1);
                if (row == null) {
                    row = sheet.createRow(rowNum - 1);
                }

                for (int col = startCol - 1; col < endCol; col++) {
                    Cell cell = row.getCell(col);
                    if (cell == null) {
                        cell = row.createCell(col);
                    }
                    cell.setCellStyle(style);
                }
            }

            try (OutputStream fos = new FileOutputStream(filePath)) {
                workbook.write(fos);
            }
        }

        return new BooleanValue(true);
    }
}
