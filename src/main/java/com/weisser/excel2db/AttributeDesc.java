package com.weisser.excel2db;

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
