package com.alesiapavliuchenkova.excel.pivot.table.dto;

import org.apache.poi.ss.usermodel.Cell;

/**
 * Created by alesia on 11/20/17.
 */
public class DataDTO {
    private Cell cell;
    private int rowNumber;
    private int columnNumber;

    public DataDTO(Cell cell, int rowNumber, int columnNumber) {
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

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;

        DataDTO dataDTO = (DataDTO) o;

        if (rowNumber != dataDTO.rowNumber) return false;
        if (columnNumber != dataDTO.columnNumber) return false;
        return cell != null ? cell.equals(dataDTO.cell) : dataDTO.cell == null;

    }

    @Override
    public int hashCode() {
        int result = cell != null ? cell.hashCode() : 0;
        result = 31 * result + rowNumber;
        result = 31 * result + columnNumber;
        return result;
    }

    @Override
    public String toString() {
        return "DataDTO{" +
                "cell=" + cell +
                ", rowNumber=" + rowNumber +
                ", columnNumber=" + columnNumber +
                '}';
    }
}
