package com.wuyue.excel.test;

import com.wuyue.excel.CellPosition;
import com.wuyue.excel.ExcelReader;
import com.wuyue.excel.POIUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressBase;

import java.io.File;
import java.io.FileInputStream;
import java.util.List;

public class Test {

    public static void main(String[] args) throws Exception {
        File file = new File("D:\\1.xlsx");
        FileInputStream fileInputStream = new FileInputStream(file);
        List<MyRow> rows = ExcelReader.builder()
                .inputStream(fileInputStream)
                .sheetNo(1)
                .headLineMun(1)
                .build()
                .read(MyRow.class);
        System.out.println(rows);
        rows.forEach(row -> {
            // 行号，如果要提示实际excel行号，应该要加上headLineMun的值
            System.out.print("Row number:" + row.getRowNum());
            // 校验结果代码(0为正常)
            System.out.print(", validate code:" + row.getValidateCode());
            // 校验结果内容
            System.out.println(", message:" + row.getValidateMessage());


            System.out.println("---------------");





        });
    }

    private CellStyle cellStyle;

    public void addComment(Workbook workbook, Sheet sheet, String address, String commentText) {

        CreationHelper factory = workbook.getCreationHelper();
        //get an existing cell or create it otherwise:
        Cell cell = POIUtils.getCell(sheet, CellPosition.of(address));

        ClientAnchor anchor = factory.createClientAnchor();
        //i found it useful to show the comment box at the bottom right corner
        anchor.setCol1(cell.getColumnIndex() + 1); //the box of the comment starts at this given column...
        anchor.setCol2(cell.getColumnIndex() + 3); //...and ends at that given column
        anchor.setRow1(cell.getRowIndex() + 1); //one row below the cell...
        anchor.setRow2(cell.getRowIndex() + 5); //...and 4 rows high

        Drawing drawing = sheet.createDrawingPatriarch();
        Comment comment = drawing.createCellComment(anchor);
        //set the comment text and author
        comment.setString(factory.createRichTextString(commentText));
        //comment.setAuthor(author);

        cell.setCellComment(comment);
        cell.setCellStyle(findCellStyle(sheet));
    }

    private CellStyle findCellStyle(Sheet sheet) {

        if (cellStyle == null) {

            Font font = sheet.getWorkbook().createFont();
            font.setFontHeightInPoints((short) 10);
            font.setFontName("맑은 고딕");

            cellStyle = sheet.getWorkbook().createCellStyle();

            cellStyle.setFont(font);

            cellStyle.setBorderRight(BorderStyle.THIN);
            cellStyle.setBorderLeft(BorderStyle.THIN);
            cellStyle.setBorderTop(BorderStyle.THIN);
            cellStyle.setBorderBottom(BorderStyle.THIN);

            cellStyle.setBottomBorderColor(IndexedColors.RED.getIndex());
            cellStyle.setTopBorderColor(IndexedColors.RED.getIndex());
            cellStyle.setLeftBorderColor(IndexedColors.RED.getIndex());
            cellStyle.setRightBorderColor(IndexedColors.RED.getIndex());

            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        }

        return cellStyle;
    }

}
