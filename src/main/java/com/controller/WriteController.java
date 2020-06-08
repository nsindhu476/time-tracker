package com.controller;

import com.dao.ExcelWrite;


import com.util.Constants;
import com.models.EmployeeConnectionDetails;
import com.models.EmployeeDetails;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;

import javax.validation.constraints.NotNull;
import javax.validation.constraints.Pattern;
import javax.validation.constraints.Size;
import java.io.File;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.List;

@Controller
public class WriteController {
    Logger logger = LoggerFactory.getLogger(WriteController.class);
    @Autowired
    ExcelWrite excelWrite;
    int count = 0;

    @GetMapping("details")
    public String logTime(@RequestParam @NotNull(message = "LID is mandatory") @Pattern(regexp =
            "^L[0-9]*$") @Size(min= 7, max = 7, message = "LID Should be valid length") String lid, @RequestParam
            String email, @RequestParam String contype) throws InvalidFormatException, IOException {
    	
        EmployeeConnectionDetails emp = new EmployeeConnectionDetails();
        emp.setLid(lid);
        emp.setEmail(email);
        emp.setContype(contype);
        
        logger.info("EmployeeConnection Object: {}", emp);
        try {
            if (fileExists()) {
      excelWrite.update(emp, lid,email,contype);
            	
            } else {
                excelWrite.write(emp);
            }
        } catch (IOException e) {
            e.printStackTrace();
            return "error";
        }
   
        return "success";
    }

 

    public boolean fileExists() {
        File file = new File(Constants.FILE_NAME1);
        logger.info("Files exists? {}", file.exists());
        return file.exists();
    }

}
