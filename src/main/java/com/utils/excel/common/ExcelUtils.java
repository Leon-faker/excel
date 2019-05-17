package com.utils.excel.common;


import com.alibaba.fastjson.JSON;
import com.utils.excel.dto.CategoryTree;
import com.utils.excel.dto.GeneratorParentId;
import com.utils.excel.dto.MergedDetails;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;

@Slf4j
public class ExcelUtils {
    private static final String EXCEL_XLS = "xls";
    private static final String EXCEL_XLSX = "xlsx";

    private static Workbook getWorkbok(InputStream in, File file) throws IOException {
        Workbook wb = null;
        if(file.getName().endsWith(EXCEL_XLS)){
            wb = new HSSFWorkbook(in);
        }else if(file.getName().endsWith(EXCEL_XLSX)){
            wb = new XSSFWorkbook(in);
        }
        return wb;
    }

    private static void checkExcelVaild(File file) throws Exception{
        if(!file.exists()){
            throw new Exception("文件不存在");
        }
        if(!(file.isFile() && (file.getName().endsWith(EXCEL_XLS) || file.getName().endsWith(EXCEL_XLSX)))){
            throw new Exception("文件不是EXCEL");
        }
    }

    private static Object getValue(Cell cell, FormulaEvaluator formulaEvaluator){
        Object obj = null;
        switch (formulaEvaluator.evaluate(cell).getCellTypeEnum()){
            case BOOLEAN:
                obj = cell.getBooleanCellValue();
                break;
            case ERROR:
                obj = cell.getErrorCellValue();
                break;
            case NUMERIC:
                obj = cell.getNumericCellValue();
                break;
            case STRING:
                obj = cell.getStringCellValue();
                break;
            default:
                break;
        }

        return obj;
    }

    public static List<CategoryTree> getExcelData(String file){
        SimpleDateFormat fmt = new SimpleDateFormat("yyyy-MM-dd");
        try {
            // 同时支持Excel 2003、2007
            File excelFile = new File(file); // 创建文件对象
            if(!excelFile.exists()){
                return null;
            }
            FileInputStream in = new FileInputStream(excelFile); // 文件流
            Workbook workbook = null;
            if(excelFile.getName().endsWith(EXCEL_XLS)){
                workbook = getWorkbok(in,excelFile);
            }else
            {
                workbook = WorkbookFactory.create(in); // 这种方式 Excel2003/2007/2010都是可以处理的
            }
            //Workbook workbook = WorkbookFactory.create(is); // 这种方式 Excel2003/2007/2010都是可以处理的
            FormulaEvaluator formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();

            int sheetCount = workbook.getNumberOfSheets(); // Sheet的数量
            /**
             * 设置当前excel中sheet的下标：0开始
             */
           Sheet sheet = workbook.getSheetAt(0);   // 遍历第一个Sheet
//            Sheet sheet = workbook.getSheetAt(2);   // 遍历第三个Sheet

            List<CategoryTree> categoryTreelst = new ArrayList<>();
            CategoryTree treeNode = null;
            CategoryTree node = null;
            //获取总行数
//          System.out.println(sheet.getLastRowNum());

            // 为跳过第一行目录设置count
            int count = 0;
            for (Row row : sheet) {
                try {
                    // 跳过第一和第二行的目录
                    if(count < 1 ) {
                        count++;
                        continue;
                    }

                    //如果当前行没有数据，跳出循环
//                    if(row.getCell(0).toString().equals("")){
//                        System.out.println("跳出");
//                        return;
//                    }

                    //获取总列数(空格的不计算)
                    int columnTotalNum = row.getPhysicalNumberOfCells();
//                    System.out.println("总列数：" + columnTotalNum);
//
//                    System.out.println("最大列数：" + row.getLastCellNum());

                    //for循环的，不扫描空格的列
//                    for (Cell cell : row) {
//                    	System.out.println(cell);
//                    }
                    MergedDetails mergedDetails = isMergedRegion(sheet, row.getCell(0));
                    int end = row.getLastCellNum();
                    for (int i = 0; i < end; i++) {

                        Cell cell = row.getCell(i);
                       // System.out.println(cell.getRowIndex()+1+":"+mergedDetails.startRow);
                        if(i == 0){
                            if(cell.getRowIndex()+1 == mergedDetails.getStartRow()){
                                treeNode = new CategoryTree();
                            }
                        }
                        if(cell == null || cell.toString().equals("")) {
                            //System.out.println("null" + "\t");
                            continue;
                        }
                        Object obj = getValue(cell,formulaEvaluator);

                        if(i == 0){
                            treeNode.setName(String.valueOf(getValue(cell,formulaEvaluator)));
                            categoryTreelst.add(treeNode);
                        }
                        if(i == 1){
                            node = treeNode.addTreeNode(String.valueOf(getValue(cell,formulaEvaluator)));
                            node.setParentId(treeNode.getId()); // 添加parentId
                        }
                        if(i == 2){
                            node.addTreeNode(String.valueOf(getValue(cell,formulaEvaluator))).setParentId(node.getId()); // 添加parentId
                        }
                        //System.out.print(obj + "\t");
                        //System.out.println(cell.getRowIndex());
                       //System.out.println(JSON.toJSONString(mergedDetails));
                    }
                } catch (Exception e) {
                    log.error("解析excel格式异常："+e.getMessage());
                    e.printStackTrace();
                }
                //System.out.println(JSON.toJSONString(categoryTreelst));
            }
            // 重置Id
            new GeneratorParentId().ResetId();
            log.info("解析结果："+JSON.toJSONString(categoryTreelst));
            return categoryTreelst;
        } catch (Exception e) {
            log.error("excel文件异常："+e.getMessage());
            e.printStackTrace();
        }
        return null;
    }

    private static MergedDetails isMergedRegion(Sheet sheet,Cell cell){
        for(int i = 0;i<sheet.getNumMergedRegions();i++){
            CellRangeAddress ca = sheet.getMergedRegion(i);
            int firstColumn = ca.getFirstColumn();
            int lastColumn = ca.getLastColumn();
            int firstRow = ca.getFirstRow();
            int lastRow = ca.getLastRow();
            if(cell.getRowIndex() >= firstRow && cell.getRowIndex() <= lastRow){
                if(cell.getColumnIndex() >= firstColumn && cell.getColumnIndex() <= firstColumn){
                    return new MergedDetails(true,firstRow+1,lastRow+1,firstColumn+1,lastColumn+1);
                }
            }
        }
        return new MergedDetails(false,0,0,0,0);
    }
}
