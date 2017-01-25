import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;

/**
 * A little tool to import data from an excel file into a JDBC database.
 *
 * Sheet -> tables
 * Row1 -> Attributes names
 * Row2 -> Attribute types
 *
 * TODO Two modes: DIRECT MODE (writes directly) and SQL MODE (just generate SQL statements to stdout)
 * TODO Usage info
 * DONE Generate TABLE Drop Statements if requested.
 * DONE Collect Metadata.
 * TODO Make Insert attribute writes more safe by using Prepared Statements.
 *
 * Created by q186379 on 24.01.2017.
 */
public class Excel2DB {
    private boolean dropTablesFirst = true;

    // JDBC database connection
    private Connection connection;

    // Table Descriptions (filled on the fly)
    List<TableDesc> tableDescList;

    public static void main(String[] args) {
        // String filename = args[1];
        String filename = null;

        if (args.length >= 1) {
            filename = args[0];
        } else {
            filename = "c:/portal_m_dev/workspace/excel2db/src/main/resources/sample.xlsx";
        }
        new Excel2DB().run(filename);
    }

    public Excel2DB() {
        tableDescList = new ArrayList<>();
    }

    public void run(String filename) {
        run(filename, "org.h2.Driver", "jdbc:h2:~/test2", "sa", null);
    }

    public void run(String filename, String jdbcDriverClassName, String jdbcUrl, String jdbcUsername, String jdbcPassword) {
        // Connect to database
        connectDB(jdbcDriverClassName, jdbcUrl, jdbcUsername, jdbcPassword);

        // Open file
        try {
            InputStream inp = new FileInputStream(filename);

            XSSFWorkbook wb = null;
            try {
                wb = new XSSFWorkbook(inp);

                // Get list of sheet names. This will be our table names.
                for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                    StringBuffer sqlCreate = new StringBuffer();

                    // Only process visible sheets. If invisible, just skip.
                    if (wb.isSheetHidden(i)) continue;

                    Sheet sheet = wb.getSheetAt(i);
                    final String tablename = sheet.getSheetName();

                    // Add metadata
                    TableDesc tabledesc = new TableDesc(tablename);
                    tableDescList.add(tabledesc);

                    // TODO Drop table first
                    if (dropTablesFirst) {
                        StringBuffer sqlDrop = new StringBuffer();
                        sqlCreate.append("DROP TABLE IF EXISTS ");
                        sqlCreate.append(tablename);
                        sqlCreate.append(";\n");

                        // Print and execute on database
                        System.out.println(sqlDrop.toString());
                        executeUpdate(sqlDrop.toString());
                    }

                    sqlCreate.append("CREATE TABLE ");
                    sqlCreate.append(tablename);
                    sqlCreate.append(" ( ");

                    // Generate schema

                    // Get the column names from row a of each sheet
                    // Get the datatypes from second column
                    Row attributeRow = sheet.getRow(0);
                    Row dataTypeRow = sheet.getRow(1);

                    if (attributeRow != null && dataTypeRow != null) {
                        final int maxColumns = attributeRow.getLastCellNum();
                        for (int col = 0; col < maxColumns; col++) {
                            Cell cell1 = attributeRow.getCell(col);
                            Cell cell2 = dataTypeRow.getCell(col);

                            if (cell1 != null && cell2 != null) {
                                String attributeName = cell1.getStringCellValue();
                                String dataTypeName = cell2.getStringCellValue();

                                sqlCreate.append("\t");

                                sqlCreate.append("\"");
                                sqlCreate.append(attributeName);
                                sqlCreate.append("\"");

                                sqlCreate.append(" ");
                                sqlCreate.append(dataTypeName);

                                // Add to metadata
                                DataType dataType = DataType.getDataTypeForName(dataTypeName);
                                tabledesc.addAttribute(new AttributeDesc(attributeName, dataType));

                                // If datatype in Excel is BOLD -> Primary key
                                if (isBold(wb, cell2)) {
                                    sqlCreate.append(" PRIMARY KEY");
                                }

                                // Do not append behind last attribute
                                if (col < maxColumns - 1) {
                                    sqlCreate.append(", ");
                                }
                            }
                        }
                    }
                    sqlCreate.append(");\n");


                    // Print and execute on database
                    System.out.println(sqlCreate.toString());
                    executeUpdate(sqlCreate.toString());

                    // Generate data
                    final int maxRows = sheet.getLastRowNum();
                    for (int row = 2; row <= maxRows; row++) {
                        Row currentRow = sheet.getRow(row);

                        if (currentRow != null) {
                            StringBuffer sqlInsert = new StringBuffer();

                            sqlInsert.append("INSERT INTO ");
                            sqlInsert.append(tablename);
                            sqlInsert.append(" VALUES (");

                            final int maxColumns = attributeRow.getLastCellNum();
                            for (int col = 0; col < maxColumns; col++) {
                                Cell cell = currentRow.getCell(col);

                                DataType cellTypeMeta = tabledesc.get(col).getDataType();

                                if (cell != null) {
                                    processCell(sqlInsert, cell, cellTypeMeta);
                                } else {
                                    processCellEmpty(sqlInsert, cellTypeMeta);

                                    // Warning
                                    //System.err.println("NULL cell at row: " + row + " column: " + col);
                                }

                                if (col < maxColumns - 1) {
                                    sqlInsert.append(", ");
                                }
                            }

                            sqlInsert.append(");");

                            // Print and execute on database
                            System.out.println(sqlInsert.toString());
                            executeUpdate(sqlInsert.toString());
                        }
                    }
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }

        // Bye bye database
        disconnectDB();
    }

    private void connectDB(final String driverClass, final String jdbcUrl, final String username, final String password) {
        try {
            Class.forName(driverClass);

            connection = DriverManager.getConnection(jdbcUrl, username, password);
        } catch(ClassNotFoundException e) {
            e.printStackTrace();
            System.out.println("Error: unable to load driver class!");
        } catch (SQLException e) {
            e.printStackTrace();
        }
    }

    private void disconnectDB() {
        if (connection != null) {
            try {
                if (!connection.isClosed()) {
                    connection.close();
                    connection = null;
                }
            } catch (SQLException e) {
                e.printStackTrace();
            }
        }
    }

    private void executeUpdate(final String sql) {
        Statement stmt = null;
        try {
            stmt = connection.createStatement();
            stmt.executeUpdate(sql);
        } catch (SQLException e ) {
            e.printStackTrace();
        } finally {
            if (stmt != null) {
                try {
                    stmt.close();
                } catch (SQLException e ) {
                    e.printStackTrace();
                }
            }
        }
    }


    private boolean isBold(Workbook wb, Cell cell) {
        boolean isBold = false;

        XSSFCellStyle style = (XSSFCellStyle) cell.getCellStyle();
        XSSFFont font = style.getFont();
        isBold = font.getBold();

        return isBold;
    }

    /**
     * Escape string data before placing it into SQL Insert
     * @param sql
     * @return
     */
    private String escapeSQLAttribute(String sql) {
        String result = sql.replace("'", "''");
        return result;
    }

    // Mapping:
    // DataType    CellType    Action
    // ----------------------------------
    // INTEGER     NUMERIC     as is
    // DATE        NUMERIC     getCellDateValue()
    private void processCell(StringBuffer sqlInsert, Cell cell, DataType cellTypeMeta) {
        CellType type = cell.getCellTypeEnum();

        if (type == CellType.NUMERIC && cellTypeMeta == DataType.INTEGER) {
            double dataNum = cell.getNumericCellValue();
            sqlInsert.append(dataNum);
        } else if (type == CellType.NUMERIC && cellTypeMeta == DataType.VARCHAR2) {
            double dataNum = cell.getNumericCellValue();
            sqlInsert.append("'");
            sqlInsert.append(escapeSQLAttribute(Double.toString(dataNum));
            sqlInsert.append("'");

        } else if (type == CellType.STRING && cellTypeMeta == DataType.INTEGER) {
            String dataString = cell.getStringCellValue();
            sqlInsert.append(dataString);
        } else if (type == CellType.STRING && cellTypeMeta == DataType.VARCHAR2) {
            String dataString = cell.getStringCellValue();
            sqlInsert.append("'");
            sqlInsert.append(escapeSQLAttribute(dataString));
            sqlInsert.append("'");
        } else if (type == CellType.NUMERIC && cellTypeMeta == DataType.DATE) {
            DateFormat df = new SimpleDateFormat("dd MM yyyy");
            java.util.Date date = cell.getDateCellValue();

            // PARSEDATETIME('26 Jul 2016, 05:15:58 AM','dd MMM yyyy, hh:mm:ss a','en')

            // H2 Syntax
            sqlInsert.append("PARSEDATETIME(");
            sqlInsert.append("'");
            sqlInsert.append(df.format(date));
            sqlInsert.append("',");
            sqlInsert.append("'dd MM yyyy','en')");
        }

    }

    private void processCellEmpty(StringBuffer sqlInsert, DataType cellTypeMeta) {
        if (cellTypeMeta == DataType.INTEGER) {
            sqlInsert.append("null");
        } else if (cellTypeMeta == DataType.VARCHAR2) {
            sqlInsert.append("''");
        } else if (cellTypeMeta == DataType.DATE) {
            sqlInsert.append("null");
        }
    }
}
