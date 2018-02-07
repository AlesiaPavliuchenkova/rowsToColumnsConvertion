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
    private long start;
    private long readTime;
    private long pivotTime;
    private long writeTime;
    private ExecutorService executorService;

    public ConcurrentWorkBookHandler(ExecutorService executorService) {
        this.executorService = executorService;
    }

    public void readDataFile(File file, DataWorkBook dataWorkBook) throws IOException {
        start = System.currentTimeMillis();
        try (FileInputStream in = new FileInputStream(file)) {
            XSSFWorkbook workbook = new XSSFWorkbook(in);
            workbook.forEach((sheet) -> createDataSheet(workbook, sheet, dataWorkBook));
        }
        readTime = System.currentTimeMillis() - start;
        System.out.println(String.format("Read  concurrently time: %s ms", readTime));
    }

    public DataWorkBook pivotFileData(DataWorkBook dataWorkBook) {
        start = System.currentTimeMillis();
        DataWorkBook pivotDataWorkBook = new DataWorkBook();
        dataWorkBook.getDataSheetList().forEach((sheet) -> createPivotSheet(sheet, pivotDataWorkBook));

        pivotTime = System.currentTimeMillis() - start;
        System.out.println(String.format("Pivot concurrently time: %s ms", pivotTime));
        return pivotDataWorkBook;

    }

    public void writeDataToFile(File file, DataWorkBook dataWorkbook) throws IOException, InvalidFormatException {
        start = System.currentTimeMillis();
        try (FileOutputStream out = new FileOutputStream(file, true)) {
            Workbook workbook = WorkbookFactory.create(file);
            dataWorkbook.getDataSheetList().forEach((sh) -> createSheet(workbook, sh));
            workbook.write(out);
            writeTime = System.currentTimeMillis() - start;
            System.out.println(String.format("Write concurrently time: %s ms", writeTime));
            System.out.println(String.format("Full concurrent process time = %s ms", readTime + pivotTime + writeTime));
        }
    }

    private void createDataSheet(Workbook workbook, Sheet sheet, DataWorkBook dataWorkBook) {
        DataSheet dataSheet = new DataSheet(workbook.getSheetIndex(sheet.getSheetName()));
        List<Future> futures = new ArrayList<>();
        sheet.forEach((currentRow) -> futures.add(executorService.submit(
                () -> createDataRow(currentRow, dataSheet))));
        futures.forEach((future) -> {
            try {
                if (!future.isDone()) future.get();
            } catch (InterruptedException ex) {
                ex.printStackTrace();
            } catch (ExecutionException ex) {
                ex.printStackTrace();
            }
        });
        dataWorkBook.addSheet(dataSheet);
    }

    private void createDataRow(Row currentRow, DataSheet dataSheet) {
        DataRow dataRow = new DataRow(currentRow.getRowNum());
        currentRow.forEach((cell) -> dataRow.addData(new DataDTO(cell, cell.getRowIndex(), cell.getColumnIndex())));
        dataSheet.addRow(dataRow);
    }

    private void createPivotSheet(DataSheet sheet, DataWorkBook pivotDataWorkBook) {
        DataSheet pivotSheet = new DataSheet(sheet.getSheetNumber());
        sheet.getDataRowList().forEach((row) -> {
            try {
                row.getDataList()
                        .forEach((dataDTO -> pivotSheet(dataDTO, pivotSheet)));
            } catch (Exception ex) {
                ex.printStackTrace();
            }
        });
        pivotDataWorkBook.addSheet(pivotSheet);
    }

    private void pivotSheet(DataDTO dataDTO, DataSheet pivotSheet) {
        DataHandler dataHandler = new DataHandler(dataDTO);
        dataHandler.switchRowColumn();
        if (pivotSheet.getDataRowList().size() - 1 < dataDTO.getRowNumber()) {
            pivotSheet.addRow(new DataRow(dataDTO.getRowNumber()));
        }
        pivotSheet.getRow(dataDTO.getRowNumber()).addData(dataDTO);
    }

    private void createSheet(Workbook workbook, DataSheet dataSheet) {
        XSSFSheet sheet = (XSSFSheet) workbook.createSheet("Sheet" + (workbook.getNumberOfSheets() + 1));
        dataSheet.getDataRowList().forEach((dataRow) -> createRow(sheet, dataRow));
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
}