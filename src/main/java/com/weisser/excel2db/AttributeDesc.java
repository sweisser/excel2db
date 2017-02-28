package com.weisser.excel2db;

/**
 * Created by q186379 on 25.01.2017.
 */
public class AttributeDesc {
    private String name;
    private DataType dataType;

    public AttributeDesc(String name, DataType dataType) {
        this.name = name;
        this.dataType = dataType;
    }

    public String getName() {
        return name;
    }

    public DataType getDataType() {
        return dataType;
    }
}
