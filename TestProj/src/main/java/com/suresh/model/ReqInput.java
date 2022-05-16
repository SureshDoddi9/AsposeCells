package com.suresh.model;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class ReqInput {
    private String fileName;
    private String sheetName;
    private List<String> colsToUpdate;
    private String functionality;
}
