package com.suresh.service;

import com.aspose.cells.*;

import java.io.File;
import java.io.FileInputStream;

public class FormatRow {

    private static final String FILE_NAME = "E:/MyProj/test1.xlsx";
    private String updatedFile = "E:/MyProj/updated2.xlsx";


    public void  formatRow() throws Exception {
        FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
        Workbook workbook = new Workbook(excelFile);

        Style style = workbook.createStyle();
        Font font = style.getFont();
        font.setSize(12);
        font.setBold(true);
        font.setColor(Color.getGreen());



        style.setVerticalAlignment(TextAlignmentType.CENTER);

// Setting the horizontal alignment of the text in the "A1" cell
        style.setHorizontalAlignment(TextAlignmentType.CENTER);
        style.setPattern(BackgroundType.SOLID);
        style.setForegroundColor(Color.getYellow());

        StyleFlag styleFlag = new StyleFlag();
        styleFlag.setFont(true);
        styleFlag.setVerticalAlignment(true);
        styleFlag.setHorizontalAlignment(true);
        styleFlag.setCellShading(true);




    //    Row row = workbook.getWorksheets().get(0).getCells().getRows().getRowByIndex(0);
// Assigning the Style object to the Style property of the row

        Range range = workbook.getWorksheets().get(0).getCells().createRange("A1","C1");

        ColumnCollection columns = workbook.getWorksheets().get(0).getCells().getColumns();



        range.applyStyle(style, styleFlag);


// Saving the Excel file
        workbook.save(updatedFile);

    }
}
