/**
 * Created by q186379 on 25.01.2017.
 */
public enum DataType {
    INTEGER, VARCHAR2, DATE;

    public static DataType getDataTypeForName(String name) {
        switch (name) {
            case "INTEGER":
                return INTEGER;
            case "VARCHAR2":
                return VARCHAR2;
            case "DATE":
                return DATE;
            default:
                return null;
        }
    }
}

