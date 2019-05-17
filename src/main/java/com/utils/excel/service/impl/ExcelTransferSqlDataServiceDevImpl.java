package com.utils.excel.service.impl;

import com.alibaba.fastjson.JSON;
import com.utils.excel.common.ExcelUtils;
import com.utils.excel.dto.CategoryTree;
import com.utils.excel.mapper.categoryMapper;
import com.utils.excel.mapper.pojo.category;
import com.utils.excel.mapper.pojo.categoryExample;
import com.utils.excel.service.ExcelTransferSqlDataService;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.springframework.stereotype.Service;
import org.springframework.util.CollectionUtils;

import javax.annotation.Resource;
import java.io.*;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

@Service
@Slf4j
public class ExcelTransferSqlDataServiceDevImpl implements ExcelTransferSqlDataService {

    @Resource
    public categoryMapper categoryMapperDto;

 //   public static void main(String[] args) throws IOException {
//        File file = new File("C://shop1.xls");
//
//        String[][] result = getData(file, 1);
//
//        int rowLength = result.length;
//
//        for(int i=0;i<rowLength;i++) {
//
//            for(int j=0;j<result[i].length;j++) {

//                System.out.print(result[i][j]+"\t\t");
//                String value = result[i][j];
//            }
//
//            System.out.println();
//
//        }
//        String c = "C://shop.xls";
//        System.out.println(JSON.toJSONString(c.split(".")));
 //   }

    public static void insert(List<CategoryTree> categoryTree){
        if(categoryTree != null){
            for(int i=0;i<categoryTree.size();i++){
                CategoryTree treeNode = categoryTree.get(i);
                if(treeNode != null){
                    category c = new category();
                    c.setStyle("");
                    c.setParentId(treeNode.getParentId().shortValue());
                    c.setCatName(treeNode.getName());
//                    dao.insert();
                    System.out.println("次数");
                    insert(treeNode.getCateChildrenTreeNode());
                }
            }
        }else {
            return;
        }

    }


    @Override
    public String inserSqlData(String fileName) {

        if(fileName == null || fileName.equals("")){
            return "File can't is null";
        }


        File file = new File(fileName);
        String[][] result = null;
        try {
            result  = getData(file, 1);
        }catch (Exception e){
            return e.getMessage();
        }


        int rowLength = result.length;

        for(int i=0;i<rowLength;i++) {
            Integer index = 0;
//            System.out.println(result[i].length);
            for (int j = 0; j < result[i].length; j++) {

//                System.out.print(result[i][j] + "\t\t");
                String value = result[i][j];
                if("".equals(value)){
                    continue;
                }
               try{
                   categoryExample cateEx = new categoryExample();
                   categoryExample.Criteria cateCr = cateEx.createCriteria();
                   cateCr.andCatNameEqualTo(value);
                   cateCr.andParentIdEqualTo(index.shortValue());
                   List<category> categorylst = categoryMapperDto.selectByExample(cateEx);
                   if(CollectionUtils.isEmpty(categorylst)) {
                       category c = new category();
                       c.setCatName(value);
                       c.setParentId(index.shortValue());
                       c.setSortOrder(true);
                       c.setStyle("");
                       categoryMapperDto.insertSelective(c);
                       index = c.getCatId().intValue();
//                    System.out.println("index:"+index);
                   }else {
                       index  = categorylst.get(0).getCatId().intValue();
                   }
               }catch (Exception e){
                   log.error("插入到数据库异常:"+e.getMessage());
                   return "插入到数据库异常:"+e.getMessage();
               }
            }
            index = 0;
            System.out.println();

        }
        return "Success";
    }

    /**

     * 读取Excel的内容，第一维数组存储的是一行中格列的值，二维数组存储的是多少个行

     * @param file 读取数据的源Excel

     * @param ignoreRows 读取数据忽略的行数，比喻行头不需要读入 忽略的行数为1

     * @return 读出的Excel中数据的内容

     * @throws FileNotFoundException

     * @throws IOException

     */

    public static String[][] getData(File file, int ignoreRows)

            throws FileNotFoundException, IOException {

        List<String[]> result = new ArrayList<String[]>();

        int rowSize = 0;

        BufferedInputStream in = new BufferedInputStream(new FileInputStream(

                file));

        // 打开HSSFWorkbook

        POIFSFileSystem fs = new POIFSFileSystem(in);

        HSSFWorkbook wb = new HSSFWorkbook(fs);

        HSSFCell cell = null;

        for (int sheetIndex = 0; sheetIndex < wb.getNumberOfSheets(); sheetIndex++) {

            HSSFSheet st = wb.getSheetAt(sheetIndex);

            // 第一行为标题，不取

            for (int rowIndex = ignoreRows; rowIndex <= st.getLastRowNum(); rowIndex++) {

                HSSFRow row = st.getRow(rowIndex);

                if (row == null) {

                    continue;

                }

                int tempRowSize = row.getLastCellNum() + 1;

                if (tempRowSize > rowSize) {

                    rowSize = tempRowSize;

                }
                String[] values = new String[rowSize];

                Arrays.fill(values, "");

                boolean hasValue = false;

                for (short columnIndex = 0; columnIndex <= row.getLastCellNum(); columnIndex++) {

                    String value = "";

                    cell = row.getCell(columnIndex);
                    if (cell != null) {

                        // 注意：一定要设成这个，否则可能会出现乱码

                        // cell.setEncoding(HSSFCell.ENCODING_UTF_16);

                        switch (cell.getCellType()) {

                            case HSSFCell.CELL_TYPE_STRING:

                                value = cell.getStringCellValue();

                                break;

                            case HSSFCell.CELL_TYPE_NUMERIC:

                                if (HSSFDateUtil.isCellDateFormatted(cell)) {

                                    Date date = cell.getDateCellValue();

                                    if (date != null) {

                                        value = new SimpleDateFormat("yyyy-MM-dd")

                                                .format(date);

                                    } else {

                                        value = "";

                                    }

                                } else {

                                    value = new DecimalFormat("0").format(cell

                                            .getNumericCellValue());

                                }

                                break;

                            case HSSFCell.CELL_TYPE_FORMULA:

                                // 导入时如果为公式生成的数据则无值

                                if (!cell.getStringCellValue().equals("")) {

                                    value = cell.getStringCellValue();

                                } else {

                                    value = cell.getNumericCellValue() + "";

                                }

                                break;

                            case HSSFCell.CELL_TYPE_BLANK:

                                break;

                            case HSSFCell.CELL_TYPE_ERROR:

                                value = "";

                                break;

                            case HSSFCell.CELL_TYPE_BOOLEAN:

                                value = (cell.getBooleanCellValue() == true ? "Y"

                                        : "N");

                                break;

                            default:

                                value = "";

                        }

                    }

                    if (columnIndex == 0 && value.trim().equals("")) {

                        break;

                    }

                    values[columnIndex] = rightTrim(value);

                    hasValue = true;

                }



                if (hasValue) {

                    result.add(values);

                }

            }

        }

        in.close();
        System.out.println(JSON.toJSONString(result));

        String[][] returnArray = new String[result.size()][rowSize];

        for (int i = 0; i < returnArray.length; i++) {

            returnArray[i] = (String[]) result.get(i);

        }

        return returnArray;

    }

    /**

     * 去掉字符串右边的空格

     * @param str 要处理的字符串

     * @return 处理后的字符串

     */

    public static String rightTrim(String str) {

        if (str == null) {

            return "";

        }

        int length = str.length();

        for (int i = length - 1; i >= 0; i--) {

            if (str.charAt(i) != 0x20) {

                break;

            }

            length--;

        }

        return str.substring(0, length);

    }
}
