package com.suresh.service;

import com.aspose.cells.*;
import org.apache.tomcat.jni.Directory;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

@Service
public class AsposeService {

    private static final String FILE_NAME = "E:/MyProj/test1.xlsx";
    private String updatedFile = "E:/MyProj/updated.xlsx";
    public void setStyles() throws Exception {
// Instantiating a Workbook object

        FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
        Workbook workbook = new Workbook(excelFile);

// Accessing the first worksheet in the Excel file
        Worksheet worksheet = workbook.getWorksheets().get(0);
        Cells cells = worksheet.getCells();



// Accessing the "A1" cell from the worksheet
        RowCollection headers = cells.getRows();



// Adding some value to the "A1" cell
         Cell cell = cells.get("A1");
        cell.setValue("Hello Aspose! this is suresh");

        Style style = cell.getStyle();

// Setting the vertical alignment of the text in the "A1" cell
       style.setVerticalAlignment(TextAlignmentType.CENTER);

// Setting the horizontal alignment of the text in the "A1" cell
        style.setHorizontalAlignment(TextAlignmentType.CENTER);

// Setting the font color of the text in the "A1" cell
        Font font = style.getFont();
       font.setColor(Color.getGreen());
       font.setBold(true);
       font.setSize(15);


// Setting the cell to shrink according to the text contained in it
        style.setShrinkToFit(true);

// Setting the bottom border
        style.setBorder(BorderType.BOTTOM_BORDER, CellBorderType.MEDIUM, Color.getRed());

// Saved style
        cell.setStyle(style);

// Saving the modified Excel file in default (that is Excel 2003) format
        workbook.save(updatedFile);
        System.out.println("done............");
    }
}
