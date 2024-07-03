package com.evangsol.report.util;

import com.alibaba.excel.write.handler.SheetWriteHandler;
import com.alibaba.excel.write.handler.context.SheetWriteHandlerContext;
import org.apache.poi.ss.usermodel.*;

public class CustomStyleHandler implements SheetWriteHandler {

    @Override
    public void beforeSheetCreate(SheetWriteHandlerContext context) {
        // Do nothing
    }

    @Override
    public void afterSheetCreate(SheetWriteHandlerContext context) {
        Sheet sheet = context.getWriteSheetHolder().getSheet();
        Workbook workbook = context.getWriteWorkbookHolder().getWorkbook();

        for (Row row : sheet) {
            for (Cell cell : row) {
                CellStyle originalStyle = cell.getCellStyle();
                if (originalStyle != null) {
                    CellStyle newStyle = workbook.createCellStyle();
                    newStyle.cloneStyleFrom(originalStyle);
                    cell.setCellStyle(newStyle);
                }
            }
        }
    }
}
