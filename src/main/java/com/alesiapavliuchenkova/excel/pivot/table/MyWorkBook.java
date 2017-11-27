package com.alesiapavliuchenkova.excel.pivot.table;

import java.util.ArrayList;

/**
 * Created by alesia on 11/22/17.
 */
public class MyWorkBook {
    private ArrayList<MySheet> mySheets;

    public MyWorkBook() {
        this.mySheets = new ArrayList<>();
    }

    public MyWorkBook(ArrayList<MySheet> mySheets) {
        this.mySheets = mySheets;
    }

    public void addSheet(MySheet mySheet) {
        mySheets.add(mySheet);
    }

    public ArrayList<MySheet> getMySheets() {
        return mySheets;
    }

    public void setMySheets(ArrayList<MySheet> mySheets) {
        this.mySheets = mySheets;
    }
}
