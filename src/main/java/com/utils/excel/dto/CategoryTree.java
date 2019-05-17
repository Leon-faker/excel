package com.utils.excel.dto;


import java.util.ArrayList;
import java.util.List;

public class CategoryTree<T> {
    private Integer id = 0;
    private String name = "";
    private Integer parentId = 0;

    private List<CategoryTree<T>> cateChildrenTreeNode = new ArrayList<>();

    public CategoryTree(){
        //构造id
        id = new GeneratorParentId().getId();
    }

    public CategoryTree<T> addTreeNode(String name){
        CategoryTree cate = new CategoryTree();
        cate.setName(name);
        cateChildrenTreeNode.add(cate);
        return cate;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public List<CategoryTree<T>> getCateChildrenTreeNode() {
        return cateChildrenTreeNode;
    }

    public void setCateChildrenTreeNode(List<CategoryTree<T>> cateChildrenTreeNode) {
        this.cateChildrenTreeNode = cateChildrenTreeNode;
    }

    public Integer getParentId() {
        return parentId;
    }

    public void setParentId(Integer parentId) {
        this.parentId = parentId;
    }

    public Integer getId() {
        return id;
    }

    public void setId(Integer id) {
        this.id = id;
    }
}
