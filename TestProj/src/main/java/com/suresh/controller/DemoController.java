package com.suresh.controller;

import com.suresh.model.ReqInput;
import com.suresh.service.DemoService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;

@RestController
public class DemoController {

    @Autowired
    private DemoService demoService;

    @PostMapping("/bulkUpdate")
    public String excelOperations(@RequestBody(required = false) ReqInput reqInput) throws Exception {
        return demoService.updateOperations(reqInput);
    }
}
