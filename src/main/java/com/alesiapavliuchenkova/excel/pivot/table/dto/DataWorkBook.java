package com.alesiapavliuchenkova.excel.pivot.table.dto;

import java.util.ArrayList;

/**
 * Created by alesia on 11/22/17.
 */
public class DataWorkBook {
    private ArrayList<DataSheet> dataSheetList;

    public DataWorkBook() {
        this.dataSheetList = new ArrayList<>();
    }

    public DataWorkBook(ArrayList<DataSheet> dataSheets) {
        this.dataSheetList = dataSheets;
    }

    public void addSheet(DataSheet dataSheet) {
        dataSheetList.add(dataSheet);
    }

    public ArrayList<DataSheet> getDataSheetList() {
        return dataSheetList;
    }

    public void setDataSheetList(ArrayList<DataSheet> dataSheetList) {
        this.dataSheetList = dataSheetList;
    }
}
