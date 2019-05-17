package com.utils.excel.dto;

import java.util.ArrayList;
import java.util.List;

public class GeneratorParentId {
    private static Integer id = 0;
    private static List<Integer> idlst;

    public GeneratorParentId(){
        idlst = new ArrayList<>();
    }

    public Integer getId(){
        idlst.add(id);
        id++;
        return id;
    }

    public void ResetId(){
        id = 0;
    }
}
