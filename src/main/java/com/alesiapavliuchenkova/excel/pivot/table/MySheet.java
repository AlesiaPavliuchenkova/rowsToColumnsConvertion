package com.alesiapavliuchenkova.excel.pivot.table;

import java.util.ArrayList;

/**
 * Created by alesia on 11/20/17.
 */
public class MySheet {
    private ArrayList<MyRow> myRows;
    private int sheetNumber;

    public MySheet(int sheetNumber) {
        this.myRows = new ArrayList<MyRow>();
        this.sheetNumber = sheetNumber;
    }

    public void addRow(MyRow myRow) {
        myRows.add(myRow);
    }

    public ArrayList<MyRow> getMyRows() {
        return myRows;
    }

    public void setMyRows(ArrayList<MyRow> myRows) {
        this.myRows = myRows;
    }


    public int getSheetNumber() {
        return sheetNumber;
    }

    public void setSheetNumber(int sheetNumber) {
        this.sheetNumber = sheetNumber;
    }

    public MyRow getRow(int index) {
        return myRows.get(index);
    }
}
