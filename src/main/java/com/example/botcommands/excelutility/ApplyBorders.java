package com.example.botcommands.excelutility;

import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.BooleanValue;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.model.AttributeType;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
//Test
@BotCommand
@CommandPkg(
        name = "ApplyBorders",
        label = "Apply Borders to Excel Range",
        node_label = "Apply borders in {{filePath}}",
        description = "Applies borders to a specific cell range in the selected sheet",
        icon = "icons/Microsoft_Office_Excel_Logo.png"
)
public class ApplyBorders {

    @Execute
    public Value<Boolean> action(
            @Idx(index = "1", type = AttributeType.FILE)
            @Pkg(label = "Excel File Path")
            String filePath,

            @Idx(index = "2", type = AttributeType.TEXT)
            @Pkg(label = "Sheet Name")
            String sheetName,

            @Idx(index = "3", type = AttributeType.NUMBER)
            @Pkg(label = "Start Row Number")
            Double startRow,

            @Idx(index = "4", type = AttributeType.NUMBER)
            @Pkg(label = "End Row Number")
            Double endRow,

            @Idx(index = "5", type = AttributeType.NUMBER)
            @Pkg(label = "Start Column Number")
            Double startCol,

            @Idx(index = "6", type = AttributeType.NUMBER)
            @Pkg(label = "End Column Number")
            Double endCol
    ) throws Exception {

        if (filePath == null || filePath.isEmpty()) {
            throw new Exception("File path is required.");
        }

        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = (sheetName != null && !sheetName.isEmpty())
                    ? workbook.getSheet(sheetName)
                    : workbook.getSheetAt(0);

            if (sheet == null) {
                throw new Exception("Sheet not found: " + sheetName);
            }

            // Convert Double → int
            int sRow = startRow != null ? startRow.intValue() - 1 : 0;
            int eRow = endRow != null ? endRow.intValue() - 1 : sheet.getLastRowNum();
            int sCol = startCol != null ? startCol.intValue() - 1 : 0;
            int eCol = endCol != null ? endCol.intValue() - 1 : 0;

            for (int r = sRow; r <= eRow; r++) {
                Row row = sheet.getRow(r);
                if (row == null) row = sheet.createRow(r);

                for (int c = sCol; c <= eCol; c++) {
                    Cell cell = row.getCell(c);
                    if (cell == null) cell = row.createCell(c);

                    CellStyle style = workbook.createCellStyle();

                    style.setBorderTop(BorderStyle.THIN);
                    style.setBorderBottom(BorderStyle.THIN);
                    style.setBorderLeft(BorderStyle.THIN);
                    style.setBorderRight(BorderStyle.THIN);

                    style.setTopBorderColor(IndexedColors.BLACK.getIndex());
                    style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
                    style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
                    style.setRightBorderColor(IndexedColors.BLACK.getIndex());

                    cell.setCellStyle(style);
                }
            }

            try (FileOutputStream fos = new FileOutputStream(filePath)) {
                workbook.write(fos);
            }
        }

        return new BooleanValue(true);
    }
}
