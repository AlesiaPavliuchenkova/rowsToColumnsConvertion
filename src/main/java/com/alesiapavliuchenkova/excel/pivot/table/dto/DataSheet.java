package com.alesiapavliuchenkova.excel.pivot.table.dto;

import java.util.ArrayList;

/**
 * Created by alesia on 11/20/17.
 */
public class DataSheet {
    private ArrayList<DataRow> dataRowList;
    private int sheetNumber;

    public DataSheet(int sheetNumber) {
        this.dataRowList = new ArrayList<DataRow>();
        this.sheetNumber = sheetNumber;
    }

    public void addRow(DataRow dataRow) {
        dataRowList.add(dataRow);
    }

    public ArrayList<DataRow> getDataRowList() { return dataRowList; }

    public void setDataRowList(ArrayList<DataRow> rowList) {
        this.dataRowList = rowList;
    }


    public int getSheetNumber() {
        return sheetNumber;
    }

    public void setSheetNumber(int sheetNumber) {
        this.sheetNumber = sheetNumber;
    }

    public DataRow getRow(int index) {
        return dataRowList.get(index);
    }
}
