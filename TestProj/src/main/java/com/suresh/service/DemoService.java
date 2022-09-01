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
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

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
        FileInputStream excelFile = new FileInputStream(new File(ff));
        Workbook workbook = new Workbook(excelFile);

        Worksheet worksheet = workbook.getWorksheets().get("Sheet1");
        Cells cells = worksheet.getCells();

        for(int i=0;i<=cells.getMaxColumn();i++) {
            if(i==0){
                cells.get(cells.getMaxRow()+1,i).setValue("Difference");
            }else {
                double doubleSum = 0;
                int intSum = 0;
                //Amount Column is in Double Type
                if (i == 3){
                    for (int j = cells.getMaxRow() - 2; j > 0; j--) {
                        doubleSum = doubleSum + cells.get(j, i).getDoubleValue();
                    }
                    System.out.println(cells.get(cells.getMaxRow() - 1, i).getDoubleValue());
                    double totalData = cells.get(cells.getMaxRow() - 1, i).getDoubleValue() - doubleSum;
                    cells.get(cells.getMaxRow(), i).setValue(totalData);
                }else{
                    //remaining columns are in int type
                    for (int j = cells.getMaxRow() - 2; j > 0; j--) {
                        intSum = intSum + cells.get(j, i).getIntValue();
                    }
                    System.out.println(cells.get(cells.getMaxRow() - 1, i).getIntValue());
                    int totalData = cells.get(cells.getMaxRow() - 1, i).getIntValue() - intSum;
                    cells.get(cells.getMaxRow(), i).setValue(totalData);
                }
            }
        }
        workbook.save(ff2,SaveFormat.XLSX);
    }


}
