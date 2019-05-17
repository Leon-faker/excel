package com.utils.excel.mapper;

import com.utils.excel.mapper.pojo.category;
import com.utils.excel.mapper.pojo.categoryExample;
import java.util.List;
import org.apache.ibatis.annotations.Param;

public interface categoryMapper {
    int countByExample(categoryExample example);

    int deleteByExample(categoryExample example);

    int deleteByPrimaryKey(Short catId);

    int insert(category record);

    int insertSelective(category record);

    List<category> selectByExample(categoryExample example);

    category selectByPrimaryKey(Short catId);

    int updateByExampleSelective(@Param("record") category record, @Param("example") categoryExample example);

    int updateByExample(@Param("record") category record, @Param("example") categoryExample example);

    int updateByPrimaryKeySelective(category record);

    int updateByPrimaryKey(category record);
}