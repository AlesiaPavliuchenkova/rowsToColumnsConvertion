package com.alesiapavliuchenkova.excel.pivot.table;

import java.util.ArrayList;

/**
 * Created by alesia on 11/20/17.
 */
public class MyRow {
    private ArrayList<Data> datas;
    private int rowNumber;

    public MyRow(int rowNumber) {
        this.datas = new ArrayList<>();
        this.rowNumber = rowNumber;
    }

    public void addData(Data data) {
        datas.add(data);
    }

    public ArrayList<Data> getDatas() {
        return datas;
    }

    public void setDatas(ArrayList<Data> datas) {
        this.datas = datas;
    }

    public int getRowNumber() {
        return rowNumber;
    }

    public void setRowNumber(int rowNumber) {
        this.rowNumber = rowNumber;
    }
}
