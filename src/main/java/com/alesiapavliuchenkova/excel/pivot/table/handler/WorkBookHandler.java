package com.alesiapavliuchenkova.excel.pivot.table.handler;

import com.alesiapavliuchenkova.excel.pivot.table.dto.DataWorkBook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.xmlbeans.impl.piccolo.io.FileFormatException;

import java.io.File;
import java.io.IOException;

/**
 * Created by alesia on 2/3/18.
 */
public interface WorkBookHandler {

    default void validateFormat(File file) throws FileFormatException {
        String[] arr = file.toString().split("\\.");

        if(!(arr[arr.length - 1].equals("xls") || arr[arr.length - 1].equals("xlsx")))
            throw new FileFormatException("Invalid format. Format can be xls/xlsx.");
    }

    void readDataFile(File file, DataWorkBook workBook) throws IOException;
    DataWorkBook pivotFileData(DataWorkBook dataWorkBook);
    void writeDataToFile(File file, DataWorkBook dataWorkbook) throws IOException, InvalidFormatException;

}
