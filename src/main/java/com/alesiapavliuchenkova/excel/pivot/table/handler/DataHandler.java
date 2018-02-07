package com.alesiapavliuchenkova.excel.pivot.table.handler;

import com.alesiapavliuchenkova.excel.pivot.table.dto.DataDTO;

/**
 * Created by alesia on 2/2/18.
 */
public class DataHandler {
    DataDTO dataDTO;

    public DataHandler(DataDTO dataDTO) {
        this.dataDTO = dataDTO;
    }

    public void switchRowColumn() {
        int colNum = dataDTO.getColumnNumber();
        dataDTO.setColumnNumber(dataDTO.getRowNumber());
        dataDTO.setRowNumber(colNum);
    }
}
