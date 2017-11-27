package com.alesiapavliuchenkova.excel.pivot.table;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.impl.piccolo.io.FileFormatException;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

/**
 * Created by apavliuchenkova on 23/11/2017.
 */
public class MyWorkBookHandler {
    public static void validateFormat(File file) throws FileFormatException
    {
        String[] arr = file.toString().split("\\.");

        if(!(arr[arr.length - 1].equals("xls") || arr[arr.length - 1].equals("xlsx")))
            throw new FileFormatException("Invalid format. Format can be xls/xlsx.");
    }

    /***************************************/
    public static void readDataFile(File file, MyWorkBook myWorkBook) throws IOException
    { //move to MyWorkBook read method
        try(FileInputStream in = new FileInputStream(file)) {
            Workbook workbook = new XSSFWorkbook(in);
            int sheetCount = workbook.getNumberOfSheets();

            for (int i = 0; i < sheetCount; i++) {
                MySheet mySheet = new MySheet(i);

                Sheet dataTypeSheet = workbook.getSheetAt(i);
                Iterator<Row> iterator = dataTypeSheet.iterator();
                int rowNumber = 0;

                while (iterator.hasNext()) {
                    Row currentRow = iterator.next();
                    Iterator<Cell> cellIterator = currentRow.cellIterator();
                    int columnNumber = 0;
                    MyRow myRow = new MyRow(currentRow.getRowNum());

                    while (cellIterator.hasNext()) {
                        Cell cell = cellIterator.next();
                        Data data = new Data(cell, rowNumber, columnNumber);
                        myRow.addData(data);
                        columnNumber++;
                    }
                    rowNumber++;
                    mySheet.addRow(myRow);
                }
                myWorkBook.addSheet(mySheet);
            }
        }
    }

    /***************************************/
    public static MyWorkBook pivotFileData(MyWorkBook myWorkBook)
    { //change to pivot and write - then concurrency can be added
        MyWorkBook pivotWorkBook = new MyWorkBook();
        MySheet pivotSheet = null;
        for (MySheet sheet: myWorkBook.getMySheets()) {
            pivotSheet = new MySheet(sheet.getSheetNumber());
            for (MyRow row: sheet.getMyRows()) {
                for (Data data: row.getDatas()) {
                    data.switchRowColumn();
                    if(pivotSheet.getMyRows().size() - 1 < data.getRowNumber()) {
                        pivotSheet.addRow(new MyRow(data.getRowNumber()));
                    }
                    pivotSheet.getRow(data.getRowNumber()).addData(data); //key value storage
                    //write to file - I have sheet number && row Number && column Number
                }
            }
        }
        pivotWorkBook.addSheet(pivotSheet);
        return pivotWorkBook;
    }

    /***************************************/
    public static void writeDataToFile(File file, MyWorkBook myWorkbook)
            throws IOException, InvalidFormatException
    {
        try(FileOutputStream out = new FileOutputStream(file, true)) {
            Workbook workbook = WorkbookFactory.create(file);

            myWorkbook.getMySheets().forEach((sh) ->
                {
                    XSSFSheet sheet = (XSSFSheet) workbook.createSheet("Sheet" + (workbook.getNumberOfSheets() + 1));
                    sh.getMyRows().forEach((myRow) ->
                        {
                            Row row = sheet.createRow(myRow.getRowNumber());
                            myRow.getDatas().forEach((data) -> setData(row, data));
                        });
                });
            workbook.write(out);
        }
    }

    /***************************************/
    private static void setData(Row row, Data data)
    {
        int cellNum = data.getColumnNumber();
        Cell cell = row.createCell(cellNum);

        switch (data.getCell().getCellType()) {
            case Cell.CELL_TYPE_BLANK:
                cell.setCellValue("");
                break;
            case Cell.CELL_TYPE_BOOLEAN:
                cell.setCellValue(data.getCell().getBooleanCellValue());
                break;
            case Cell.CELL_TYPE_NUMERIC:
                if (DateUtil.isCellDateFormatted(data.getCell())) {
                    cell.setCellValue(data.getCell().getDateCellValue());
                } else {
                    cell.setCellValue(data.getCell().getNumericCellValue());
                }
                break;
            case Cell.CELL_TYPE_STRING:
                cell.setCellValue(data.getCell().getStringCellValue());
                break;
            case Cell.CELL_TYPE_FORMULA:
                cell.setCellValue(data.getCell().getCellFormula());
                break;
            default: break;
        }
    }
}
