package com.suresh.controller;

import com.suresh.model.ReqInput;
import com.suresh.service.DemoService;
import org.json.JSONObject;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;

import java.util.Map;

@RestController
public class DemoController {

    @Autowired
    private DemoService demoService;

    @PostMapping("/bulkUpdate")
    public String excelOperations(@RequestBody Map<String,Object> object) throws Exception {
        return demoService.updateOperations(object);
    }
}
