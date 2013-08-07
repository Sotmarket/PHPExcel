<?php
/**
 * PHPExcel_Writer_PDFHtml
 * Формирует html для PDF
 * @category   PHPExcel
 * @package	PHPExcel_Writer_HTML
 * @copyright  sabsab
 * @author sabsab
 */
class PHPExcel_Writer_PDFHtml extends PHPExcel_Writer_HTML implements PHPExcel_Writer_IWriter {
	/**
	 * Is the current writer creating PDF?
	 *
	 * @var boolean
	 */
	protected $_isPdf = true;


    /**
     * @var Pager
     */
    protected $pager ;
    protected $isSplitTables = true;
    protected $isUseCellWidth = false;


    /**
     * Разделять ли шапку от основной таблицы
     * @param $isSplitTables
     * @return $this
     */
    public function setIsSplitTables($isSplitTables)
    {
        $this->isSplitTables = $isSplitTables;
        return $this;
    }

    protected function getIsSplitTables()
    {
        return $this->isSplitTables;
    }


    /**
     * Использовать td.width
     * @param $isUseCellWidth
     * @return $this
     */
    protected function setIsUseCellWidth($isUseCellWidth)
    {
        $this->isUseCellWidth = $isUseCellWidth;
        return $this;
    }

    /**
     * @return bool
     */
    protected function getIsUseCellWidth()
    {
        return $this->isUseCellWidth;
    }

    /**
     * @param IExcelPager $pager
     * @return $this
     */
    public function setPager(IExcelPager $pager=null)
    {
        $this->pager = $pager;
        return $this;
    }

    /**
     * @return Pager
     */
    protected function getPager()
    {
        if (null == $this->pager){
            $this->pager = new PHPExcel_Tools_Pager_DefaultPager($this->_phpExcel->getActiveSheet());
        }
        return $this->pager;
    }
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
            $css['html']['font-size']		= '11pt';
            $css['html']['background-color'] = 'white';
        }


        // table { }
        $css['table']['border-collapse']  = 'collapse';
        if (!$this->_isPdf) {
            $css['table']['page-break-after'] = 'always';
        }

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
        if ($this->getIsUseCellWidth()){
            $styleTag = " td";
        }
        else {
            $styleTag = " col";
        }
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
                $css['table.sheet' . $sheetIndex . $styleTag .'.col' . $column]['width'] = '42pt';
            }

            // col elements, loop through columnDimensions and set width
            foreach ($sheet->getColumnDimensions() as $columnDimension) {
                if (($width = PHPExcel_Shared_Drawing::cellDimensionToPixels($columnDimension->getWidth(), $this->_defaultFont)) >= 0) {
                    $width = PHPExcel_Shared_Drawing::pixelsToPoints($width);
                    $column = PHPExcel_Cell::columnIndexFromString($columnDimension->getColumnIndex()) - 1;
                    $this->_columnWidths[$sheetIndex][$column] = $width;
                    $css['table.sheet' . $sheetIndex . $styleTag .'.col' . $column]['width'] = $width . 'pt';

                    if ($columnDimension->getVisible() === false) {
                        $css['table.sheet' . $sheetIndex .$styleTag. '.col' . $column]['visibility'] = 'collapse';
                        $css['table.sheet' . $sheetIndex . $styleTag.'.col' . $column]['*display'] = 'none'; // target IE6+7
                    }
                }
            }

            // Default row height
            $rowDimension = $sheet->getDefaultRowDimension();

            // table.sheetN tr { }
            $css['table.sheet' . $sheetIndex . ' tr'] = array();

            if ($rowDimension->getRowHeight() == -1) {
                $pt_height = PHPExcel_Shared_Font::getDefaultRowHeightByFont($this->_phpExcel->getDefaultStyle()->getFont());
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
     * Высота документа в поинтах
     * @param $sheet
     * @return int
     */
    protected function getDocumentWidth($sheetIndex){

        return array_sum($this->_columnWidths[$sheetIndex]);

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

        if (!$this->_useInlineCss) {
            $gridlines = $pSheet->getShowGridLines() ? ' gridlines' : '';
            $style = "";
            if ($this->getIsUseCellWidth()){
                $style = 'style="overflow: wrap;width:100% "';
            }
            else {
               // $style = 'style="width:'.($this->getDocumentWidth($sheetIndex)).'pt "';
                if (null !=$this->getTableWidth()){
                    $style = 'style="width:'.$this->getTableWidth().'"';
                }

            }
            $html .= '	<table border="0" cellpadding="0" '.$style.'  cellspacing="0"  class="sheet' . $sheetIndex . $gridlines . '">' . PHP_EOL;
        } else {
            $style = isset($this->_cssStyles['table']) ?
                $this->_assembleCSS($this->_cssStyles['table']) : '';

            if ($this->_isPdf && $pSheet->getShowGridLines()) {
                $html .= '	<table border="1" cellpadding="1" cellspacing="1"  style=" ' . $style . '">' . PHP_EOL;
            } else {
                $html .= '	<table border="0" cellpadding="1"  cellspacing="0" style="' . $style . '">' . PHP_EOL;
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

    /**
     * Построить html для одного ряда
     * @param $sheet
     * @param $row
     * @param $dimension
     * @return string
     */
    protected function buildRowHtml($sheet, $row,  $dimension){
        $html = '';
        if ( !isset($this->_isSpannedRow[$sheet->getParent()->getIndex($sheet)][$row]) ) {
            // Start a new rowData
            $rowData = array();
            // Loop through columns
            $column = $dimension[0][0] - 1;
            while($column++ < $dimension[1][0]) {
                // Cell exists?
                if ($sheet->cellExistsByColumnAndRow($column, $row)) {
                    $rowData[$column] = $sheet->getCellByColumnAndRow($column, $row);
                } else {
                    $rowData[$column] = '';
                }
            }
            $html = $this->_generateRow($sheet, $rowData, $row - 1);
        }
        return $html;

    }

    /**
     * Разделитель страниц
     * @return string
     */
    protected function getPageBreakHtml(){
        return '<div style="page-break-before:always" ></div>';
    }
	/**
	 * Generate sheet data
	 *
	 * @return	string
	 * @throws PHPExcel_Writer_Exception
	 */
	public function generateSheetData() {
		// PHPExcel object known?
		if (is_null($this->_phpExcel)) {
			throw new PHPExcel_Writer_Exception('Internal PHPExcel object not set to an instance of an object.');
		}

		// Ensure that Spans have been calculated?
		if (!$this->_spansAreCalculated) {
			$this->_calculateSpans();
		}

		// Fetch sheets
		$sheets = array();
		if (is_null($this->_sheetIndex)) {
			$sheets = $this->_phpExcel->getAllSheets();
		} else {
			$sheets[] = $this->_phpExcel->getSheet($this->_sheetIndex);
		}

		// Construct HTML
		$html = '';

		// Loop all sheets
		$sheetId = 0;
		foreach ($sheets as $sheet) {
			// Get worksheet dimension
			$dimension = explode(':', $sheet->calculateWorksheetDimension());
			$dimension[0] = PHPExcel_Cell::coordinateFromString($dimension[0]);
			$dimension[0][0] = PHPExcel_Cell::columnIndexFromString($dimension[0][0]) - 1;
			$dimension[1] = PHPExcel_Cell::coordinateFromString($dimension[1]);
			$dimension[1][0] = PHPExcel_Cell::columnIndexFromString($dimension[1][0]) - 1;

			// row min,max
			$rowMin = $dimension[0][1];
			$rowMax = $dimension[1][1];

			// calculate start of <tbody>, <thead>
			$tbodyStart = $rowMin;
			$theadStart = $theadEnd   = 0; // default: no <thead>	no </thead>

			if ($sheet->getPageSetup()->isRowsToRepeatAtTopSet()) {
				$rowsToRepeatAtTop = $sheet->getPageSetup()->getRowsToRepeatAtTop();

                $theadStart = $rowsToRepeatAtTop[0];
                $theadEnd   = $rowsToRepeatAtTop[1];
                $tbodyStart = $rowsToRepeatAtTop[1] + 1;

			}
			// Loop through cells
			$row = $rowMin-1;

            $pager = $this->getPager();
            $pages  =$pager->getSmoothedPageMap();

            $countHeaders = $theadEnd -$theadStart;
            $headerHtml="";

            for ($numPage=1; $numPage<=count($pages); $numPage++){

                $page = $pages[$numPage];
                $start = $page->getStart();
                $end = ($page->getFinish()<$rowMax)?$page->getFinish():$rowMax;
                if ($start>$rowMax){
                    continue;
                }
                $html .= $this->_generateTableHeader($sheet);
                if ($numPage == 1){
                    for ($row = $start; $row<=$end; $row++){
                        if ($row == $theadStart) {

                            if (true == $this->getIsSplitTables()){
                                $html.=$this->_generateTableFooter();
                                $html.=$this->_generateTableHeader($sheet);
                            }
                        }
                        $rowHtml = $this->buildRowHtml($sheet, $row, $dimension);
                        $html .= $rowHtml;
                        if ($row>=$theadStart && $row<=$theadEnd){
                            $headerHtml .=$rowHtml ;
                        }

                    }

                }
                else{
                    $html.=$headerHtml;
                    $tbodyStart = $start;
                    for ($row = $start; $row<=$end; $row++){

                        $rowHtml = $this->buildRowHtml($sheet, $row, $dimension);
                        $html .=$rowHtml;
                    }

                }
                $html .= $this->_extendRowsForChartsAndImages($sheet, $row);

                // Close table body.
                // Write table footer
                $html .= $this->_generateTableFooter();
                if ($numPage<count($pages)){
                    $html.=$this->getPageBreakHtml();
                }
            }
			// Writing PDF?
			if ($this->_isPdf) {
				if (is_null($this->_sheetIndex) && $sheetId + 1 < $this->_phpExcel->getSheetCount()) {
					$html.=$this->getPageBreakHtml();;
				}
			}

			// Next sheet
			++$sheetId;
		}

		// Return
		return $html;
	}
	
}
