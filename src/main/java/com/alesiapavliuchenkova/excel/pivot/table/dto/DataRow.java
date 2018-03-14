package com.alesiapavliuchenkova.excel.pivot.table.dto;

import java.util.ArrayList;
import java.util.List;

/**
 * Created by alesia on 11/20/17.
 */
public class DataRow {
    private List<DataDTO> dataList;
    private int rowNumber;

    public DataRow(int rowNumber) {
        this.dataList = new ArrayList<>();
        this.rowNumber = rowNumber;
    }


    public void addData(DataDTO dataDTO) {
        dataList.add(dataDTO);
    }

    public List<DataDTO> getDataList() {
        return dataList;
    }

    public void setDataList(ArrayList<DataDTO> dataList) {
        this.dataList = dataList;
    }

    public int getRowNumber() {
        return rowNumber;
    }

    public void setRowNumber(int rowNumber) {
        this.rowNumber = rowNumber;
    }
}
