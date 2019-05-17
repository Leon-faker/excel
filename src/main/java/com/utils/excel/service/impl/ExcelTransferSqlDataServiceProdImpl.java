package com.utils.excel.service.impl;

import com.utils.excel.common.ExcelUtils;
import com.utils.excel.dto.CategoryTree;
import com.utils.excel.mapper.categoryMapper;
import com.utils.excel.mapper.pojo.category;
import com.utils.excel.service.ExcelTransferSqlDataService;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections4.CollectionUtils;
import org.springframework.stereotype.Service;

import javax.annotation.Resource;
import java.io.File;
import java.util.List;

@Slf4j
@Service
public class ExcelTransferSqlDataServiceProdImpl implements ExcelTransferSqlDataService{

    @Resource
    private categoryMapper categoryMapperDto;

    @Override
    public String inserSqlData(String fileName) {
        List<CategoryTree> categoryTreeList = ExcelUtils.getExcelData(fileName);
        try {
            return  recursionInsertToDataBase(categoryTreeList);
        } catch (Exception e) {
            e.printStackTrace();
            return "数据插入异常";
        }
    }

    public static void main(String[] args) {
//        List<CategoryTree> categoryTreeList = ExcelUtils.getExcelData("");
//        recursionInsertToDataBase(categoryTreeList);
        File file = new File("C:/shop.xls");
        System.out.println(file.exists());
    }

    private String recursionInsertToDataBase(List<CategoryTree> categoryTreeList){
        if(CollectionUtils.isEmpty(categoryTreeList)){
            return "文件不存在";
        }
        for(CategoryTree cate : categoryTreeList){
            category cateDto = new category();
            if(null != cateDto) {
               try{
                   cateDto.setCatName(cate.getName());
                   cateDto.setParentId(cate.getParentId().shortValue());
                   cateDto.setStyle("");
                   categoryMapperDto.insertSelective(cateDto);
                   cate.setParentId(cateDto.getParentId().intValue());
                   recursionInsertToDataBase(cate.getCateChildrenTreeNode());
               }catch (Exception e){
                    log.error("插入数据失败："+e.getMessage());
                   return "插入数据失败:"+e.getMessage();
               }
            }
        }
        return "Success";
    }
}
