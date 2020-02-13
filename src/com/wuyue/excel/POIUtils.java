package com.wuyue.excel;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.awt.*;

/**
 * Apache POIとJExcel APIの差を埋めるユーティリティクラス。
 *
 * @version 2.0
 * @author T.TSUCHIE
 *
 */
public class POIUtils {


    /**
     * シートの種類を判定する。
     *
     * @since 2.0
     * @param sheet 判定対象のオブジェクト
     * @return シートの種類。不明な場合はnullを返す。
     * @throws IllegalArgumentException {@literal sheet == null}
     */
    public static SpreadsheetVersion getVersion(final Sheet sheet) {
        ArgUtils.notNull(sheet, "sheet");

        if(sheet instanceof HSSFSheet) {
            return SpreadsheetVersion.EXCEL97;

        } else if(sheet instanceof XSSFSheet) {
            return SpreadsheetVersion.EXCEL2007;
        }

        return null;
    }

    /**
     * シートの最大列数を取得する。
     * <p>{@literal jxl.Sheet.getColumns()}</p>
     * @param sheet シートオブジェクト
     * @return 最大列数
     * @throws IllegalArgumentException {@literal sheet == null.}
     */
    public static int getColumns(final Sheet sheet) {
        ArgUtils.notNull(sheet, "sheet");

        int minRowIndex = sheet.getFirstRowNum();
        int maxRowIndex = sheet.getLastRowNum();
        int maxColumnsIndex = 0;
        for(int i = minRowIndex; i <= maxRowIndex; i++) {
            final Row row = sheet.getRow(i);
            if(row == null) {
                continue;
            }

            final int column = row.getLastCellNum();
            if(column > maxColumnsIndex) {
                maxColumnsIndex = column;
            }
        }

        return maxColumnsIndex;
    }

    /**
     * シートの最大行数を取得する
     *
     * <p>{@literal jxl.Sheet.getRows()}</p>
     * @param sheet シートオブジェクト
     * @return 最大行数
     * @throws IllegalArgumentException {@literal sheet == null.}
     */
    public static int getRows(final Sheet sheet) {
        ArgUtils.notNull(sheet, "sheet");
        return sheet.getLastRowNum() + 1;
    }

    /**
     * シートから任意アドレスのセルを取得する。
     * @since 0.5
     * @param sheet シートオブジェクト
     * @param address アドレス（Point.x=column, Point.y=row）
     * @return セル
     * @throws IllegalArgumentException {@literal sheet == null or address == null.}
     */
    public static Cell getCell(final Sheet sheet, final Point address) {
        ArgUtils.notNull(sheet, "sheet");
        ArgUtils.notNull(address, "address");
        return getCell(sheet, address.x, address.y);
    }

    /**
     * シートから任意アドレスのセルを取得する。
     * @since 1.4
     * @param sheet シートオブジェクト
     * @param address セルのアドレス
     * @return セル
     * @throws IllegalArgumentException {@literal sheet == null or address == null.}
     */
    public static Cell getCell(final Sheet sheet, final CellPosition address) {
        ArgUtils.notNull(sheet, "sheet");
        ArgUtils.notNull(address, "address");
        return getCell(sheet, address.getColumn(), address.getRow());
    }

    /**
     * シートから任意アドレスのセルを取得する。
     *
     * <p>{@literal jxl.Sheet.getCell(int column, int row)}</p>
     * @param sheet シートオブジェクト
     * @param column 列番号（0から始まる）
     * @param row 行番号（0から始まる）
     * @return セル
     * @throws IllegalArgumentException {@literal sheet == null}
     */
    public static Cell getCell(final Sheet sheet, final int column, final int row) {
        ArgUtils.notNull(sheet, "sheet");

        Row rows = sheet.getRow(row);
        if(rows == null) {
            rows = sheet.createRow(row);
        }

        Cell cell = rows.getCell(column);
        if(cell == null) {
            cell = rows.createCell(column, CellType.BLANK);
        }

        return cell;
    }



}
