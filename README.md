# mdbtocsv

### Description

mdbtocsv is a Windows console application written in c#.

mdbtocsv reads the given Access file (mdb|accdb) file and exports each user table to a csv file.

#### mdbtocsv is NOT feature complete and is a work in progress.

### Prerequisites 
- 64 bit ODBC Access Driver installed.

### Usage

**mdbtocsv** -s:"sourceFileName" -noprompt -lower

The above example will extract all tables in source file as separate CSVs

**PARAMETERS**:

|OPTION|DESCRIPTION|
|----- | ----- |
|-s:sourceFileName | (required) The Access file to process. Wrap in quotes if you have spaces in the path.|
|-noprompt | disables the 'Continue' prompt. Program will run without the need for user interaction.|
|-o:directory | user specified output directory. Default = same path as source file.|
|-nolog | disables the runtime log. !!! Set as very first parameter to fully disable log.|
|-lower | force output filenames to be lowercase.|
|-p | use PIPE '\|' Delimiter in output file.|
|-t | use TAB Delimiter in output file.|
|-c | Replaces all symbol chars with '_' and converts FieldName to UPPER case|
|-debug | enables debug 'verbose' mode.|


**Example**

mdbtocsv.exe "-s:c:\data\receipts.mdb" -p -c -lower

This would parse all tables found in the receipts.mdb file, using the PIPE delimiter, cleaning the fieldnames, and writing the output filenames in lowercase.


### License and Contact
[MIT License](https://mit-license.org/)

mdbtocsv also uses the [csvhelper](https://joshclose.github.io/csvhelper/) library with its own license.

[Contact me](mailto:geoff@bentonvillebase.com) for any questions.

