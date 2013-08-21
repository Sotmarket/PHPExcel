<?php
/**
 * Html нормально разбитый на страницы
 * @category  
 * @package   
 * @subpackage 
 * @author: sabsab
 * @date: 05.08.13
 * @version    $Id: $
 */

class PHPExcel_Writer_PageableHtml extends PHPExcel_Writer_PDFHtml implements PHPExcel_Writer_IWriter {


    /**
     * Is the current writer creating PDF?
     *
     * @var boolean
     */
    protected $_isPdf = false;
    protected $isUseCellWidth=false;
    protected $isSplitTables = false;

    /**
     * Build CSS styles
     *
     * @param	boolean	$generateSurroundingHTML	Generate surrounding HTML style? (html { })
     * @return	array
     * @throws	PHPExcel_Writer_Exception
     */
    public function buildCSS($generateSurroundingHTML = true) {
        // PHPExcel object known?
        if (is_null($this->_phpExcel)) {
            throw new PHPExcel_Writer_Exception('Internal PHPExcel object not set to an instance of an object.');
        }

        // Cached?
        if (!is_null($this->_cssStyles)) {
            return $this->_cssStyles;
        }

        // Ensure that spans have been calculated
        if (!$this->_spansAreCalculated) {
            $this->_calculateSpans();
        }

        // Construct CSS
        $css = array();

        // Start styles
        if ($generateSurroundingHTML) {
            // html { }
            $css['html']['font-family']	  = 'Calibri, Arial, Helvetica, sans-serif';
            $css['html']['font-size']		= '0pt';
            $css['html']['background-color'] = 'white';
        }


        // table { }
        $css['table']['border-collapse']  = 'collapse';
        //if (!$this->_isPdf) {
        //    $css['table']['page-break-after'] = 'always';
        //}

        // .gridlines td { }
        $css['.gridlines td']['border'] = '1px dotted black';

        // .b {}
        $css['.b']['text-align'] = 'center'; // BOOL

        // .e {}
        $css['.e']['text-align'] = 'center'; // ERROR

        // .f {}
        $css['.f']['text-align'] = 'right'; // FORMULA

        // .inlineStr {}
        $css['.inlineStr']['text-align'] = 'left'; // INLINE

        // .n {}
        $css['.n']['text-align'] = 'right'; // NUMERIC

        // .s {}
        $css['.s']['text-align'] = 'left'; // STRING

        // Calculate cell style hashes
        foreach ($this->_phpExcel->getCellXfCollection() as $index => $style) {

            $css['td.style' . $index] = $this->_createCSSStyle( $style );
        }

        // Fetch sheets
        $sheets = array();
        if (is_null($this->_sheetIndex)) {
            $sheets = $this->_phpExcel->getAllSheets();
        } else {
            $sheets[] = $this->_phpExcel->getSheet($this->_sheetIndex);
        }

        // Build styles per sheet
        foreach ($sheets as $sheet) {
            // Calculate hash code
            $sheetIndex = $sheet->getParent()->getIndex($sheet);

            // Build styles
            // Calculate column widths
            $sheet->calculateColumnWidths();

            // col elements, initialize
            $highestColumnIndex = PHPExcel_Cell::columnIndexFromString($sheet->getHighestColumn()) - 1;
            $column = -1;
            while($column++ < $highestColumnIndex) {
                $this->_columnWidths[$sheetIndex][$column] = 42; // approximation
                $css['table.sheet' . $sheetIndex . ' col.col' . $column]['width'] = '42pt';
            }

            // col elements, loop through columnDimensions and set width
            foreach ($sheet->getColumnDimensions() as $columnDimension) {
                if (($width = PHPExcel_Shared_Drawing::cellDimensionToPixels($columnDimension->getWidth(), $this->_defaultFont)) >= 0) {
                    $width = PHPExcel_Shared_Drawing::pixelsToPoints($width);
                    $column = PHPExcel_Cell::columnIndexFromString($columnDimension->getColumnIndex()) - 1;
                    $this->_columnWidths[$sheetIndex][$column] = $width;
                    $css['table.sheet' . $sheetIndex . ' col.col' . $column]['width'] = $width . 'pt';

                    if ($columnDimension->getVisible() === false) {
                        $css['table.sheet' . $sheetIndex . ' col.col' . $column]['visibility'] = 'collapse';
                        $css['table.sheet' . $sheetIndex . ' col.col' . $column]['*display'] = 'none'; // target IE6+7
                    }
                }
            }

            // Default row height
            $rowDimension = $sheet->getDefaultRowDimension();

            // table.sheetN tr { }
            $css['table.sheet' . $sheetIndex . ' tr'] = array();

            if ($rowDimension->getRowHeight() == -1) {
                $pt_height = PHPExcel_Shared_Font::getDefaultRowHeightByFont($this->_defaultFont);
            } else {
                $pt_height = $rowDimension->getRowHeight();
            }
            $css['table.sheet' . $sheetIndex . ' tr']['height'] = $pt_height . 'pt';
            if ($rowDimension->getVisible() === false) {
                $css['table.sheet' . $sheetIndex . ' tr']['display']	= 'none';
                $css['table.sheet' . $sheetIndex . ' tr']['visibility'] = 'hidden';
            }

            // Calculate row heights
            foreach ($sheet->getRowDimensions() as $rowDimension) {
                $row = $rowDimension->getRowIndex() - 1;

                // table.sheetN tr.rowYYYYYY { }
                $css['table.sheet' . $sheetIndex . ' tr.row' . $row] = array();

                if ($rowDimension->getRowHeight() == -1) {
                    $pt_height = PHPExcel_Shared_Font::getDefaultRowHeightByFont($this->_phpExcel->getDefaultStyle()->getFont());
                } else {
                    $pt_height = $rowDimension->getRowHeight();
                }
                $css['table.sheet' . $sheetIndex . ' tr.row' . $row]['height'] = $pt_height . 'pt';
                if ($rowDimension->getVisible() === false) {
                    $css['table.sheet' . $sheetIndex . ' tr.row' . $row]['display'] = 'none';
                    $css['table.sheet' . $sheetIndex . ' tr.row' . $row]['visibility'] = 'hidden';
                }
            }
        }

        // Cache
        if (is_null($this->_cssStyles)) {
            $this->_cssStyles = $css;
        }

        // Return
        return $css;
    }
    /**
     * Generate table header
     *
     * @param	PHPExcel_Worksheet	$pSheet		The worksheet for the table we are writing
     * @return	string
     * @throws	PHPExcel_Writer_Exception
     */
    protected function _generateTableHeader($pSheet) {
        $sheetIndex = $pSheet->getParent()->getIndex($pSheet);

        // Construct HTML
        $html = '';
        $html .= $this->_setMargins($pSheet);

        if (!$this->_useInlineCss) {
            $gridlines = $pSheet->getShowGridLines() ? ' gridlines' : '';
            $style = "";
            if ($this->getIsUseCellWidth()){
                $style = 'style="overflow: wrap;width:100% "';
            }
            else {
                // $style = 'style="width:'.($this->getDocumentWidth($sheetIndex)).'pt "';
                if (null !=$this->getTableWidth()){
                    $width = $this->getTableWidth();
                    $style = 'style="overflow: wrap;table-layout: fixed;width:'.$width.'"';
                }

            }
            $html .= '	<table border="0" cellpadding="0" '.$style.'  cellspacing="0"  class="sheet' . $sheetIndex . $gridlines . '">' . PHP_EOL;
        } else {
            $style = isset($this->_cssStyles['table']) ?
                $this->_assembleCSS($this->_cssStyles['table']) : '';

            if ($this->_isPdf && $pSheet->getShowGridLines()) {
                $html .= '	<table border="1" cellpadding="1" id="sheet' . $sheetIndex . '" cellspacing="1" style="' . $style . '">' . PHP_EOL;
            } else {
                $html .= '	<table border="0" cellpadding="1" id="sheet' . $sheetIndex . '" cellspacing="0" style="' . $style . '">' . PHP_EOL;
            }
        }

        // Write <col> elements
        $highestColumnIndex = PHPExcel_Cell::columnIndexFromString($pSheet->getHighestColumn()) - 1;
        $i = -1;
        while($i++ < $highestColumnIndex) {
            if (!$this->_isPdf) {
                if (!$this->_useInlineCss) {
                    $html .= '		<col class="col' . $i . '">' . PHP_EOL;
                } else {
                    $style = isset($this->_cssStyles['table.sheet' . $sheetIndex . ' col.col' . $i]) ?
                        $this->_assembleCSS($this->_cssStyles['table.sheet' . $sheetIndex . ' col.col' . $i]) : '';
                    $html .= '		<col style="' . $style . '">' . PHP_EOL;
                }
            }
        }

        // Return
        return $html;
    }

}
