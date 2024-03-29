package com.suresh.service;

import com.aspose.cells.*;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.suresh.model.Functionality;
import com.suresh.model.RenameReq;
import com.suresh.model.ReqInput;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

@Service
@Slf4j
public class DemoService {
    private String FILE_NAME = "C:/Users/Dell/Downloads/filename";
    private String updatedFile = "C:/Users/Dell/Downloads/updated.xlsx";

    private String ff = "E:/Excels/Book1.xlsx";

    private String ff2 = "E:/Excels/Excels2/updatedBook1.xlsx";





    public String updateOperations(Map<String, Object> object) throws Exception {
        ObjectMapper mapper = new ObjectMapper();
        if(object.get("functionality").equals(Functionality.NUMBER_TO_TEXT.getFunction())){
            ReqInput reqObject = mapper.convertValue(object,ReqInput.class);
            numberToText(reqObject);
        }
        if(object.get("functionality").equals(Functionality.TEXT_TO_NUMBER.getFunction())){
            ReqInput reqObject = mapper.convertValue(object,ReqInput.class);
            textToNumber(reqObject);
        }
        if(object.get("functionality").equals(Functionality.RENAME_SHEET.getFunction())){
            RenameReq reqObject = mapper.convertValue(object, RenameReq.class);
            renameWorkSheet(reqObject);
        }
        if(object.get("functionality").equals(Functionality.SHEET_FORMAT.getFunction())){
            RenameReq reqObject = mapper.convertValue(object, RenameReq.class);
            changingFormat(reqObject);
        }
        return "success";
    }


    public void numberToText(ReqInput input) throws Exception {
        String updatedFileName  = FILE_NAME.replace("filename",input.getFileName());
        FileInputStream excelFile = new FileInputStream(new File(updatedFileName));
        Workbook workbook = new Workbook(excelFile);

        Worksheet worksheet = workbook.getWorksheets().get(input.getSheetName());
        Cells cells = worksheet.getCells();
        try {
            input.getColsToUpdate().forEach(col -> {
                //to get column index from column Header Name
                int column = CellsHelper.columnNameToIndex(col);
                for (int r = 1; r < cells.getMaxRow() + 1; r++) {
                    Cell cell = cells.get(r, column);
                    log.info("cell value: "+cell.getValue());
                    if (cell != null) {
                        if (cell.getType() == CellValueType.IS_NUMERIC) {
                            int data = cell.getIntValue();
                            cell.setValue(String.valueOf(data));
                            log.info("converted cell value: "+cell.getValue());
                        }
                    }
                }
            });
        }catch (Exception e){
            log.info("columns updation Failed............");
            log.error("error : "+e);
        }
        log.info("columns updated successfully");
        workbook.save(updatedFile);
    }

    public void textToNumber(ReqInput input) throws Exception {
        String updatedFileName  = FILE_NAME.replace("filename",input.getFileName());
        FileInputStream excelFile = new FileInputStream(new File(updatedFileName));
        Workbook workbook = new Workbook(excelFile);

        Worksheet worksheet = workbook.getWorksheets().get(input.getSheetName());
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

    public void renameWorkSheet(RenameReq input) throws Exception {

        String updatedFileName  = FILE_NAME.replace("filename",input.getFileName());
        FileInputStream excelFile = new FileInputStream(new File(updatedFileName));
        Workbook workbook = new Workbook(excelFile);

        Worksheet worksheet = workbook.getWorksheets().get(input.getSheetName());
        worksheet.setName(input.getRequestedName());
        workbook.save(updatedFileName);

    }
    public void changingFormat(RenameReq input) throws Exception {

        String updatedFileName  = FILE_NAME.replace("filename",input.getFileName());
        FileInputStream file = new FileInputStream(new File(updatedFileName));
        LoadOptions loadOptions=new LoadOptions(FileFormatType.CSV);
        Workbook workbook = new Workbook(file,loadOptions);

        String newFileName= input.getFileName().replace(".csv",".xlsx");
        String newFileDirectory=FILE_NAME.replace("filename",newFileName);

        workbook.save(newFileDirectory, SaveFormat.XLSX);

    }

    public void calculate() throws Exception {
        String maintab = "HRCEU";
        FileInputStream excelFile = new FileInputStream(new File(ff));
        Workbook workbook = new Workbook(excelFile);

        Worksheet worksheet = workbook.getWorksheets().get("Sheet1");
//        worksheet.getCells().deleteBlankRows();
//        worksheet.getCells().deleteBlankColumns();
        Cells cells = worksheet.getCells();
        FindOptions findoptions = new  FindOptions();
        CellArea cellArea = new CellArea();
        cellArea.StartColumn=0;
        cellArea.EndColumn = 0;
        cellArea.StartRow = 0;
        cellArea.EndRow = cells.getMaxRow();
        findoptions.setRange(cellArea);
        findoptions.setLookAtType(LookAtType.ENTIRE_CONTENT);
        Cell cell = cells.find(maintab,null,findoptions);
        int mainRow = cell.getRow();
        cell.getColumn();

        for(int col=0;col<=cells.getMaxColumn();col++) {
            if(col==0){
                int mrow = cells.getMaxRow();
                cells.get(cells.getMaxRow()+1,col).setValue("Difference");
            }else {

//                double doubleSum = 0;
//                int intSum = 0;
//                //Amount Column is in Double Type
//                if (i == 3){
//                    System.out.println("......................");
//                    for (int j = cells.getMaxRow()-1; j > 0; j--) {
//                        if(j!=mainRow) {
//                            System.out.println(cells.get(j, i).getDoubleValue());
//                            doubleSum = doubleSum + cells.get(j, i).getDoubleValue();
//                        }
//                    }
//                    System.out.println("doubleSum: "+doubleSum);
//                    System.out.println("......................");
//                    System.out.println(cells.get(mainRow, i).getDoubleValue());
//                    double totalData = cells.get(mainRow, i).getDoubleValue() - doubleSum;
//                    cells.get(cells.getMaxRow(), i).setValue(totalData);
//                }else{
//                    //remaining columns are in int type
//                    System.out.println(".....................");
//                    for (int j = cells.getMaxRow() - 1; j > 0; j--) {
//                        if(j!=mainRow) {
//                            System.out.println(cells.get(j, i).getIntValue());
//                            intSum = intSum + cells.get(j, i).getIntValue();
//                        }
//                    }
//                    System.out.println("int sum: "+intSum);
//                    System.out.println("......................");
//                    System.out.println(cells.get(mainRow, i).getIntValue());
//                    int totalData = cells.get(mainRow, i).getIntValue() - intSum;
//                    System.out.println("totaldata: "+totalData);
//                    cells.get(cells.getMaxRow(), i).setValue(totalData);
//                    System.out.println("data::::"+cells.get(cells.getMaxRow(), i).getIntValue());
//                }


                var formula = "%s%s-SUM(%s)";
                //columns
                    String subFormula = "";
                    String finalFormula = "";
                    if(col ==1) {
                      subFormula =  getMainString(cells,mainRow,"B");
                      finalFormula = String.format(formula,"B",mainRow+1,subFormula);
                      cells.get(cells.getMaxRow(),col).setFormula(finalFormula);
                      System.out.println(cells.get(cells.getMaxRow(),col));
                    }
                    if(col==2){
                        subFormula = getMainString(cells,mainRow,"C");
                        finalFormula = String.format(formula,"C",mainRow+1,subFormula);
                        cells.get(cells.getMaxRow(),col).setFormula(finalFormula);
                    }
                    if(col==3){
                        subFormula = getMainString(cells,mainRow,"D");
                        finalFormula = String.format(formula,"D",mainRow+1,subFormula);
                        cells.get(cells.getMaxRow(),col).setFormula(finalFormula);
                    }
                }

            }
         workbook.calculateFormula();
         workbook.save(ff2,SaveFormat.XLSX);
         workbook.dispose();
        }

    public String getMainString(Cells cells,int mainRow,String col){
        StringBuilder mainStr = new StringBuilder("");
        for (int j = 2; j <=cells.getMaxRow(); j++) {
                if (j != mainRow + 1) {
                    mainStr = mainStr.append("," + col + j);
                }
        }
        mainStr.replace(0,1,"");
        return mainStr.toString();
    }

    public void testmethod(){
        List<String> dates = new ArrayList<>();
        dates.add("10/5/2022");
        dates.add("7/5/2022");
        dates.add("11/5/2021");

        Collections.sort(dates, new Comparator<String>() {
            DateFormat f = new SimpleDateFormat("dd/mm/yyyy");
            @Override
            public int compare(String o1, String o2) {
                try {
                    return f.parse(o1).compareTo(f.parse(o2));
                } catch (ParseException e) {
                    throw new RuntimeException(e);
                }
            }
        });
        System.out.println(dates.toString());
    }

}
