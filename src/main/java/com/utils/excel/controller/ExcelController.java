package com.utils.excel.controller;

import com.utils.excel.service.ExcelTransferSqlDataService;
import com.utils.excel.service.impl.ExcelTransferSqlDataServiceProdImpl;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;

import javax.annotation.Resource;

@Slf4j
@Controller
@RequestMapping(value = "/excel")
public class ExcelController {

    @Resource
    public ExcelTransferSqlDataServiceProdImpl excelTransferSqlDataServiceProd;

    @GetMapping(value = "/generator")
    @ResponseBody
    public String transferSqlData(@RequestParam(value = "file") String file){
        log.info("请求参数:"+file);
        if(null == file || "".equals(file)){
            return "文件不能为空";
        }
        return excelTransferSqlDataServiceProd.inserSqlData(file);
    }
}
