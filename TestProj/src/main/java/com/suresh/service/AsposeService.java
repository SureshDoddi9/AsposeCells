package com.suresh.service;

import com.aspose.cells.*;
import org.apache.tomcat.jni.Directory;
import org.springframework.stereotype.Service;

import java.io.*;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.time.Instant;
import java.util.Date;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

@Service
public class AsposeService {

    private static final String FILE_NAME = "E:/MyProj/test1_ADD_P202201.xls";
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

    public void extractDate() throws Exception {
//        Path path = Paths.get(FILE_NAME).getFileName().toString();
//        Path file = path.getFileName();
        System.out.println(Paths.get(FILE_NAME).getFileName().toString());

        String file = Paths.get(FILE_NAME).getFileName().toString();

        String strPattern = "\\d{4}\\d{2}";

        Pattern pattern = Pattern.compile(strPattern);
        Matcher matcher = pattern.matcher(file);
        String str = null;
        while( matcher.find() ) {
            str = matcher.group();
        }
        System.out.println(str);

        SimpleDateFormat sdf = new SimpleDateFormat("yyyyMM");
        Date date1 = sdf.parse(str);
        Date date2 = sdf.parse(sdf.format(new Date()));

        System.out.println("date1 : " + sdf.format(date1));
        System.out.println("date2 : " + sdf.format(date2));

        if (date1.equals(date2)) {
            System.out.println("Date1 is equal Date2");
        }

        if (date1.after(date2)) {
            System.out.println("Date1 is after Date2");
        }

        if (date1.before(date2)) {
            System.out.println("Date1 is before Date2");
        }
    }

    public void checkFileType(){
        StringBuilder finalColumn = new StringBuilder("USD");
        System.out.println(finalColumn.toString());
//        String file = Paths.get(filePath).getFileName().toString();
//        if(file.contains("TR") || file.contains("CP")) {
//            finalColumn = new StringBuilder("CAD");
//        }
        finalColumn = new StringBuilder("CAD");
        System.out.println(finalColumn.toString());
    }
}
