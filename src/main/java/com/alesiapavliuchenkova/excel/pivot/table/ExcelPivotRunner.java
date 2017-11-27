package com.alesiapavliuchenkova.excel.pivot.table;

import org.apache.poi.POIXMLException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.xmlbeans.impl.piccolo.io.FileFormatException;
import java.io.*;
import java.util.Scanner;
import static com.alesiapavliuchenkova.excel.pivot.table.MyWorkBookHandler.*;

/**
 * Created by alesia on 11/20/17.
 */
public class ExcelPivotRunner {

    public static void main (String[] args) {
        String path = "src\\main\\resources\\Sample-Sales-Data.xlsx";

        Scanner scanner = new Scanner(System.in);
        System.out.println("Do you want to pivot default excel file? Y/N");
        String answer = scanner.next();
        if(!answer.toUpperCase().equals("Y")) {
            System.out.println("Please provide file path with \\ delimiter:");
            path = scanner.nextLine();
        }

        String[] filePaths = path.split("\\\\");
        String filePath = "";
        for (int i = 0; i< filePaths.length; i++) {
            filePath += filePaths[i];
            if (i != filePaths.length - 1) filePath += File.separator;
        }

        File file = new File(filePath);

        if(file.exists()) {//only then create MyWorkBook
            //validate that file format is xls or xlsx
            try {
                validateFormat(file); //move this validation to MyWorkBook constructor ???

                MyWorkBook myWorkBook = new MyWorkBook();
                //read excel file data
                readDataFile(file, myWorkBook); //xlsx   + add for xls
                //pivot rows to columns
                MyWorkBook pivotMyWorkBook = pivotFileData(myWorkBook);
                //writeDataToNewSheet
                writeDataToFile(file, pivotMyWorkBook);
            } catch (FileFormatException ex) {
                ex.printStackTrace();
            } catch (FileNotFoundException ex) {
                System.out.println("There is no file " + file.toString()); //add log
                ex.printStackTrace();
            } catch (IOException ex) {
                ex.printStackTrace();        //add log
            } catch (POIXMLException ex) {
                String[] arr = file.getAbsoluteFile().toString().split(File.separator);
                String fileName = arr[arr.length - 1];
                System.out.println(fileName + " has not valid data. "); //add log
            } catch (InvalidFormatException e) {
                e.printStackTrace();
            }
        } else {
            System.out.println("file " + file + " doesn't exist.");
        }
    }
}
