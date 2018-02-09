package com.alesiapavliuchenkova.excel.pivot.table.handler;

import com.alesiapavliuchenkova.excel.pivot.table.dto.*;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.*;

/**
 * Created by alesia on 2/2/18.
 */
public class ConcurrentWorkBookHandler implements WorkBookHandler {
    private ExecutorService executorService;

    public ConcurrentWorkBookHandler(ExecutorService executorService) {
        this.executorService = executorService;
    }

    public void readDataFile(File file, DataWorkBook dataWorkBook) throws IOException {
        try (FileInputStream in = new FileInputStream(file)) {
            XSSFWorkbook workbook = new XSSFWorkbook(in);
            workbook.forEach((sheet) -> createDataSheet(workbook, sheet, dataWorkBook));
        }
    }

    public DataWorkBook pivotFileData(DataWorkBook dataWorkBook) {
        DataWorkBook pivotDataWorkBook = new DataWorkBook();
        dataWorkBook.getDataSheetList().forEach((sheet) -> pivotSheet(sheet, pivotDataWorkBook));
        return pivotDataWorkBook;
    }

    public void writeDataToFile(File file, DataWorkBook dataWorkbook) throws IOException, InvalidFormatException {
        try (FileOutputStream out = new FileOutputStream(file, true)) {
            Workbook workbook = WorkbookFactory.create(file);
            dataWorkbook.getDataSheetList().forEach((sh) -> createSheet(workbook, sh));
            workbook.write(out);
        }
    }

    private void createDataSheet(Workbook workbook, Sheet sheet, DataWorkBook dataWorkBook) {
        DataSheet dataSheet = new DataSheet(workbook.getSheetIndex(sheet.getSheetName())
                , new CopyOnWriteArrayList<>());
        List<Future> futures = new ArrayList<>();
        sheet.forEach((currentRow) ->
                futures.add(executorService.submit(() -> createDataRow(currentRow, dataSheet))));
        proceedFutures(futures);
        dataWorkBook.addSheet(dataSheet);
    }

    private void createDataRow(Row currentRow, DataSheet dataSheet) {
        DataRow dataRow = new DataRow(currentRow.getRowNum());
        currentRow.forEach((cell) -> dataRow.addData(new DataDTO(cell, cell.getRowIndex(), cell.getColumnIndex())));
        dataSheet.addRow(dataRow);
    }

    private void pivotSheet(DataSheet sheet, DataWorkBook pivotDataWorkBook) {
        DataSheet pivotSheet = new DataSheet(sheet.getSheetNumber(), new CopyOnWriteArrayList<>());
        sheet.getDataRowList().forEach((row) ->
            row.getDataList().forEach(dataDTO -> pivotData(dataDTO, pivotSheet)));
        pivotDataWorkBook.addSheet(pivotSheet);
    }

    private void pivotData(DataDTO dataDTO, DataSheet pivotSheet) {
        DataHandler dataHandler = new DataHandler(dataDTO);
        dataHandler.switchRowColumn();
        if (pivotSheet.getDataRowList().size() - 1 < dataDTO.getRowNumber())
            pivotSheet.addRow(new DataRow(dataDTO.getRowNumber()));
        pivotSheet.getRow(dataDTO.getRowNumber()).addData(dataDTO);
    }

    private void createSheet(Workbook workbook, DataSheet dataSheet) {
        ArrayList<Future> futures = new ArrayList<>();
        XSSFSheet sheet = (XSSFSheet) workbook.createSheet("Sheet" + (workbook.getNumberOfSheets() + 1));
        dataSheet.getDataRowList().forEach((dataRow) ->
                futures.add(executorService.submit(() -> createRow(sheet, dataRow))));
        proceedFutures(futures);
    }

    private void createRow(XSSFSheet sheet, DataRow dataRow) {
        Row row = sheet.createRow(dataRow.getRowNumber());
        dataRow.getDataList().forEach((dataDTO) -> setData(row, dataDTO));
    }

    private void setData(Row row, DataDTO dataDTO) {
        int cellNum = dataDTO.getColumnNumber();
        Cell cell = row.createCell(cellNum);

        switch (dataDTO.getCell().getCellType()) {
            case Cell.CELL_TYPE_BLANK:
                cell.setCellValue("");
                break;
            case Cell.CELL_TYPE_BOOLEAN:
                cell.setCellValue(dataDTO.getCell().getBooleanCellValue());
                break;
            case Cell.CELL_TYPE_NUMERIC:
                if (DateUtil.isCellDateFormatted(dataDTO.getCell())) {
                    cell.setCellValue(dataDTO.getCell().getDateCellValue());
                } else {
                    cell.setCellValue(dataDTO.getCell().getNumericCellValue());
                }
                break;
            case Cell.CELL_TYPE_STRING:
                cell.setCellValue(dataDTO.getCell().getStringCellValue());
                break;
            case Cell.CELL_TYPE_FORMULA:
                cell.setCellValue(dataDTO.getCell().getCellFormula());
                break;
            default:
                break;
        }
    }

    private void proceedFutures(List<Future> futures) {
        futures.forEach(future -> {
            if (!future.isDone()) try {
                future.get();
            } catch (InterruptedException | ExecutionException e) {
                e.printStackTrace();
            }
        });
    }
}