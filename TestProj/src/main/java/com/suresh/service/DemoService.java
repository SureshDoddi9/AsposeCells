package com.suresh.service;

import com.aspose.cells.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.ArrayList;
import java.util.List;

public class DemoService {
    private static final String FILE_NAME = "E:/MyProj/test1.xlsx";
    private String updatedFile = "E:/MyProj/updated.xlsx";


    public void numberToText() throws Exception {
        FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
        Workbook workbook = new Workbook(excelFile);

        Worksheet worksheet = workbook.getWorksheets().get(0);
        Cells cells = worksheet.getCells();

        List<String> cols = new ArrayList<>();
        cols.add("A");
        cols.add("C");

        cols.forEach(col->{
            //to get column index from column Header Name
            int column = CellsHelper.columnNameToIndex(col);
            for(int r = 1;r<cells.getMaxRow()+1;r++){
                Cell cell = cells.get(r,column);
                if(cell!=null){
                    if(cell.getType() == CellValueType.IS_NUMERIC){
                        int data = cell.getIntValue();
                        cell.setValue(String.valueOf(data));
                    }
                }
            }
        });
        workbook.save(updatedFile);
    }

    public void textToNumber() throws Exception {
        FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
        Workbook workbook = new Workbook(excelFile);

        Worksheet worksheet = workbook.getWorksheets().get(0);
        Cells cells = worksheet.getCells();

        int column = CellsHelper.columnNameToIndex("D");

        for(int r = 1;r<cells.getMaxRow()+1;r++){
            Cell cell = cells.get(r,column);
            if(cell!=null){
                if(cell.getType() == CellValueType.IS_STRING){
                    String  data = cell.getStringValue().replace("S","");
                    cell.setValue(Integer.valueOf(data));
                }
            }
        }
        workbook.save(updatedFile);
    }
}
