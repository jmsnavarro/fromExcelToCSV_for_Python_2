# fromExcelToCSV_for_Python_2
Version of fromExcelToCSV tool for Python 3 now targeted to Python 2 for legacy systems.

A terminal script written in Python that reads Excel rows and export to a pipe-delimited csv file.

Developed using Visual Studio Code with Python extension.

## Source file requirements
1. Source must be placed on the same directory as the python script
2. To modify default filename, update `SRC_FILENAME` variable
3. First sheet name must be a month name (e.g. January)
4. Year must be placed at `E3` cell (e.g. 2018)
5. Type of menu must be placed at `A` cell on any row (e.g. International)
6. Data rows are read from `A` to `E` cells where `A` cell values must be numeric

### Additional source file requirements
- Will only accept *.xls Excel files and not *.xlsx (Excel 2007 or later)

### How to run
```
$ python fromExcelToCSV_py243.py
```

### Output: csv and log files
- YYYYMMSRC_FILENAME.csv
(e.g. 201804FOOD_MENU.csv)

- YYYYMMDD_HHMMHH_fromExcelToCSV_py243.py.log
(e.g. 20180402_080012_fromExcelToCSV_py243.py.log)