package com.weisser.excel2db;

import java.util.ArrayList;
import java.util.List;

/**
 * Created by q186379 on 25.01.2017.
 */
public class TableDesc {
    private String tablename;

    private List<AttributeDesc> attributes;

    public TableDesc(String tablename) {
        this.attributes = new ArrayList<>();
        this.tablename = tablename;
    }

    public void addAttribute(AttributeDesc attributeDesc) {
        attributes.add(attributeDesc);
    }

    public AttributeDesc get(int i) {
        return attributes.get(i);
    }
}
