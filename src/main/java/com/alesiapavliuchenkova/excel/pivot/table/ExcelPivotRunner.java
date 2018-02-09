package com.alesiapavliuchenkova.excel.pivot.table;

import com.alesiapavliuchenkova.excel.pivot.table.dto.DataWorkBook;
import com.alesiapavliuchenkova.excel.pivot.table.handler.ConcurrentWorkBookHandler;
import com.alesiapavliuchenkova.excel.pivot.table.handler.NonConcurrentWorkBookHandler;
import com.alesiapavliuchenkova.excel.pivot.table.handler.WorkBookHandler;
import org.apache.poi.POIXMLException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.xmlbeans.impl.piccolo.io.FileFormatException;
import java.io.*;
import java.util.Arrays;
import java.util.Scanner;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

/**
 * Created by alesia on 11/20/17.
 */
public class ExcelPivotRunner {

    private static String filePath = "";
    private static ExecutorService executorService;

    public static void main (String[] args) {
        provideFilePath();
        File file = new File(filePath);
        long start = System.currentTimeMillis();
        try {
            printSystemInfo();
            WorkBookHandler workBookHandler = (args.length > 0 && args[0].equals("-c")) ?
                    new ConcurrentWorkBookHandler(executorService = Executors.newFixedThreadPool(Integer.parseInt(args[1]))) :
                    new NonConcurrentWorkBookHandler();
            workBookHandler.validateFormat(file);
            DataWorkBook dataWorkBook = new DataWorkBook();
            workBookHandler.readDataFile(file, dataWorkBook);
            DataWorkBook pivotDataWorkBook = workBookHandler.pivotFileData(dataWorkBook);
            workBookHandler.writeDataToFile(file, pivotDataWorkBook);
        } catch (FileFormatException | InvalidFormatException ex) {
            ex.printStackTrace();
        } catch (FileNotFoundException ex) {
            System.out.println(String.format("There is no file %s", filePath)); //add log
            ex.printStackTrace();
        } catch (IOException ex) {
            ex.printStackTrace();        //add log
        } catch (POIXMLException ex) {
            System.out.println(String.format("%s has not valid data.", filePath)); //add log
        } finally {
            if (executorService != null) {
                executorService.shutdown();
                while(!executorService.isTerminated()){}
            }
        }
        System.out.println(String.format("%s process time = %s ms"
                , (args.length > 0 && args[0].equals("-c")) ? "Concurrent" : "Non concurrent"
                , System.currentTimeMillis() - start));
    }

    private static void printSystemInfo() {
        System.out.println(String.format("System info: %s %s %s",
                System.getProperty("os.name"),
                System.getProperty("os.version"),
                System.getProperty("os.arch")));
    }

    private static void provideFilePath() {
        String path = "src\\main\\resources\\Sample-Sales-Data.xlsx";

        Scanner scanner = new Scanner(System.in);
        System.out.println("Do you want to pivot default excel file? Y/N");
        String answer = scanner.next();
        if(!answer.toUpperCase().equals("Y")) {
            System.out.println("Please provide file path with \\ delimiter:");
            path = scanner.nextLine();
        }

        String[] filePaths = path.split("\\\\");
        Arrays.stream(filePaths).forEach((el) -> filePath += el + File.separator);
        filePath = filePath.replaceFirst(".$", "");
    }
}
