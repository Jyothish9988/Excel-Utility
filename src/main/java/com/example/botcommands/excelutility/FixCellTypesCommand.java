package com.example.botcommands.excelutility;

import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.commandsdk.model.*;
import com.automationanywhere.botcommand.data.impl.StringValue;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.*;
import java.util.*;

@BotCommand
@CommandPkg(
        name = "fixCellTypes",
        label = "Fix Excel Data Types",
        node_label = "Fix data types in {{filePath}}",
        description = "Convert specific Excel columns to proper data types (Number, Date, etc.)",
        icon = "icons/Microsoft_Office_Excel_Logo.png"
)
public class FixCellTypesCommand {

    @Execute
    public Value<String> action(
            @Idx(index = "1", type = AttributeType.FILE)
            @Pkg(label = "Excel File Path")
            String filePath,

            @Idx(index = "2", type = AttributeType.TEXT)
            @Pkg(label = "Sheet Name (leave blank for all sheets)")
            String sheetName,

            @Idx(index = "3", type = AttributeType.TEXT)
            @Pkg(label = "Columns to process (A,C,D or 1,3,4)")
            String columnList,

            @Idx(index = "4", type = AttributeType.BOOLEAN)
            @Pkg(label = "Apply Number Conversion", default_value = "true", default_value_type = DataType.BOOLEAN)
            Boolean fixNumbers,

            @Idx(index = "5", type = AttributeType.BOOLEAN)
            @Pkg(label = "Apply Date Conversion", default_value = "true", default_value_type = DataType.BOOLEAN)
            Boolean fixDates,

            @Idx(index = "6", type = AttributeType.FILE)
            @Pkg(label = "Output File Path (leave blank to overwrite input)")
            String outputFile
    ) {
        if (filePath == null || filePath.trim().isEmpty()) {
            return new StringValue("Error: Excel file path cannot be empty.");
        }

        File inputFile = new File(filePath);
        if (!inputFile.exists()) {
            return new StringValue("Error: File does not exist: " + filePath);
        }

        Workbook workbook = null;
        try {
            // Read workbook first, then close input stream
            try (FileInputStream fis = new FileInputStream(inputFile)) {
                workbook = new XSSFWorkbook(fis);
            }

            List<Integer> targetCols = parseColumns(columnList);
            DataFormatter formatter = new DataFormatter();

            // Create shared styles to avoid limit issues
            CellStyle numberStyle = null;
            CellStyle dateStyle = null;

            if (fixNumbers) {
                numberStyle = workbook.createCellStyle();
                numberStyle.setDataFormat(workbook.createDataFormat().getFormat("#,##0.00"));
            }

            if (fixDates) {
                dateStyle = workbook.createCellStyle();
                dateStyle.setDataFormat(workbook.createDataFormat().getFormat("dd/mmm/yy"));
            }

            List<Sheet> sheets = new ArrayList<>();
            if (sheetName == null || sheetName.trim().isEmpty()) {
                for (Sheet s : workbook) sheets.add(s);
            } else {
                Sheet s = workbook.getSheet(sheetName);
                if (s == null) {
                    workbook.close();
                    return new StringValue("Error: Sheet '" + sheetName + "' not found.");
                }
                sheets.add(s);
            }

            int modifiedCells = 0;
            for (Sheet sheet : sheets) {
                for (Row row : sheet) {
                    for (Cell cell : row) {
                        if (cell == null || cell.getCellType() == CellType.BLANK) continue;

                        int colIdx = cell.getColumnIndex();
                        if (!targetCols.isEmpty() && !targetCols.contains(colIdx)) continue;

                        String value = formatter.formatCellValue(cell).trim();
                        if (value.isEmpty()) continue;

                        boolean cellModified = false;

                        if (fixNumbers && !cellModified) {
                            cellModified = tryConvertToNumber(cell, value, numberStyle);
                        }

                        if (fixDates && !cellModified) {
                            cellModified = tryConvertToDate(cell, value, dateStyle);
                        }

                        if (cellModified) modifiedCells++;
                    }
                }
            }

            String savePath = (outputFile == null || outputFile.trim().isEmpty())
                    ? filePath
                    : outputFile;

            // Handle overwrite case by using temporary file
            if (savePath.equals(filePath)) {
                File tempFile = new File(filePath + ".tmp");
                try (FileOutputStream fos = new FileOutputStream(tempFile)) {
                    workbook.write(fos);
                }
                workbook.close();

                // Replace original file
                if (!inputFile.delete()) {
                    throw new IOException("Could not delete original file");
                }
                if (!tempFile.renameTo(inputFile)) {
                    throw new IOException("Could not rename temporary file");
                }
            } else {
                try (FileOutputStream fos = new FileOutputStream(savePath)) {
                    workbook.write(fos);
                }
                workbook.close();
            }

            return new StringValue("Successfully processed " + modifiedCells + " cells. Output: " + savePath);

        } catch (Exception e) {
            // Ensure workbook is closed even on error
            if (workbook != null) {
                try { workbook.close(); } catch (Exception ex) { }
            }
            return new StringValue("Error: " + e.getMessage());
        }
    }

    private boolean tryConvertToNumber(Cell cell, String value, CellStyle numberStyle) {
        try {
            // Remove common number formatting characters
            String cleanValue = value.replace(",", "").replace("$", "").replace(" ", "").replace("%", "");
            double num = Double.parseDouble(cleanValue);
            cell.setCellValue(num);
            cell.setCellStyle(numberStyle);
            return true;
        } catch (NumberFormatException ignored) {
            return false;
        }
    }

    private boolean tryConvertToDate(Cell cell, String value, CellStyle dateStyle) {
        Date parsedDate = tryParseDate(value);
        if (parsedDate != null) {
            cell.setCellValue(parsedDate);
            cell.setCellStyle(dateStyle);
            return true;
        }
        return false;
    }

    // Helper to parse column letters/numbers
    private List<Integer> parseColumns(String colList) {
        List<Integer> cols = new ArrayList<>();
        if (colList == null || colList.trim().isEmpty()) return cols;

        for (String c : colList.split(",")) {
            c = c.trim();
            if (c.isEmpty()) continue;

            try {
                cols.add(Integer.parseInt(c) - 1);
            } catch (NumberFormatException e) {
                cols.add(letterToIndex(c));
            }
        }
        return cols;
    }

    private int letterToIndex(String col) {
        col = col.toUpperCase().trim();
        int result = 0;
        for (char ch : col.toCharArray()) {
            if (ch < 'A' || ch > 'Z') {
                throw new IllegalArgumentException("Invalid column character: " + ch);
            }
            result = result * 26 + (ch - 'A' + 1);
        }
        return result - 1;
    }

    private Date tryParseDate(String val) {
        String[] formats = {
                "dd/MM/yyyy", "MM/dd/yyyy", "dd-MM-yyyy", "yyyy-MM-dd",
                "dd/MM/yy", "MM/dd/yy", "dd-MM-yy", "yy-MM-dd",
                "yyyy/MM/dd", "dd.MM.yyyy", "MM.dd.yyyy"
        };

        for (String fmt : formats) {
            try {
                SimpleDateFormat sdf = new SimpleDateFormat(fmt);
                sdf.setLenient(false); // Strict parsing
                return sdf.parse(val);
            } catch (ParseException ignored) {}
        }
        return null;
    }
}