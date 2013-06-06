# Fork original PHPExcel for best PDF creation

## File Formats supported

### Reading
 * BIFF 5-8 (.xls) Excel 95 and above
 * Office Open XML (.xlsx) Excel 2007 and above
 * SpreadsheetML (.xml) Excel 2003
 * Open Document Format/OASIS (.ods)
 * Gnumeric
 * HTML
 * SYLK
 * CSV

### Writing
 * BIFF 8 (.xls) Excel 95 and above
 * Office Open XML (.xlsx) Excel 2007 and above
 * HTML
 * CSV
 * PDF (using either the tcPDF, DomPDF or mPDF libraries, which need to be installed separately)


## Requirements
 * PHP version 5.2.0 or higher
 * PHP extension php_zip enabled (required if you need PHPExcel to handle .xlsx .ods or .gnumeric files)
 * PHP extension php_xml enabled
 * PHP extension php_gd2 enabled (optional, but required for exact column width autocalculation)


## Want to contribute?
Fork us!

## License
PHPExcel is licensed under [LGPL (GNU LESSER GENERAL PUBLIC LICENSE)](https://github.com/PHPOffice/PHPExcel/blob/master/license.md)
