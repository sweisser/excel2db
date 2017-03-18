# excel2db
Import data from excel tables into jdbc SQL databases.

This tool is meant to quickly import tabular data from excel into any JDBC database.

Every sheet in the excel will result in a new table in the database.

Every column in a sheet represents an attribute in that table.

The first two rows in the excel are preserved for description of attribute names and attribute types.


Currently only a few datatypes and target database dialects are supported.
