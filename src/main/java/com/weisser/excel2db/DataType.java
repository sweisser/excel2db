package com.weisser.excel2db;

/**
 * Created by q186379 on 25.01.2017.
 */
public enum DataType {
    INTEGER, VARCHAR2, DATE, BIGINT, DECIMAL, H2_INTEGER, H2_VARCHAR2, H2_DATE, H2_BIGINT, H2DECIMAL;

    public static DataType getDataTypeForName(String name) {
        switch (name) {
            case "INTEGER":
                return INTEGER;
            case "VARCHAR2":
                return VARCHAR2;
            case "DATE":
                return DATE;
            case "BIGINT":
                return BIGINT;
            case "DECIMAL":
                return DECIMAL;
            default:
                return null;
        }
    }

    public static DataType getDataTypeForNameAndDialect(String name, int dialect) {
        switch (name) {
            case "INTEGER":
                return INTEGER;
            case "VARCHAR2":
                return VARCHAR2;
            case "DATE":
                return DATE;
            case "BIGINT":
                return BIGINT;
            case "DECIMAL":
                return DECIMAL;
            default:
                return null;
        }
    }
}
