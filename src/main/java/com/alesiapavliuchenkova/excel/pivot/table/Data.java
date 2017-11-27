package com.alesiapavliuchenkova.excel.pivot.table;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;

/**
 * Created by alesia on 11/20/17.
 */
public class Data {
    private Cell cell;
    private int rowNumber;
    private int columnNumber;

    public Data(Cell cell, int rowNumber, int columnNumber) {
        this.cell = cell;
        this.rowNumber = rowNumber;
        this.columnNumber = columnNumber;
    }

    public int getRowNumber() {
        return rowNumber;
    }

    public void setRowNumber(int rowNumber) {
        this.rowNumber = rowNumber;
    }

    public int getColumnNumber() {
        return columnNumber;
    }

    public void setColumnNumber(int columnNumber) {
        this.columnNumber = columnNumber;
    }

    public Cell getCell() { return cell; }

    public void setCell(Cell cell) { this.cell = cell; }

    public void switchRowColumn() {
        int colNum = this.columnNumber;
        this.columnNumber = this.rowNumber;
        this.rowNumber = colNum;
    }
}
