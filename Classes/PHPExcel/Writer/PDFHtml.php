<?php
/**
 * PHPExcel_Writer_PDFHtml
 * ��������� html ��� PDF
 * @category   PHPExcel
 * @package	PHPExcel_Writer_HTML
 * @copyright  Copyright (c) 2006 - 2012 PHPExcel (http://www.codeplex.com/PHPExcel)
 */
class PHPExcel_Writer_PDFHtml extends PHPExcel_Writer_HTML implements PHPExcel_Writer_IWriter {


	/**
	 * Sheet index to write
	 *
	 * @var int
	 */
	private $_sheetIndex	= 0;

	/**
	 * Images root
	 *
	 * @var string
	 */
	private $_imagesRoot	= '.';

	/**
	 * embed images, or link to images
	 *
	 * @var boolean
	 */
	private $_embedImages	= false;

	/**
	 * Use inline CSS?
	 *
	 * @var boolean
	 */
	private $_useInlineCss = false;

	/**
	 * Array of CSS styles
	 *
	 * @var array
	 */
	private $_cssStyles = null;

	/**
	 * Array of column widths in points
	 *
	 * @var array
	 */
	private $_columnWidths = null;



	/**
	 * Flag whether spans have been calculated
	 *
	 * @var boolean
	 */
	private $_spansAreCalculated	= false;

	/**
	 * Excel cells that should not be written as HTML cells
	 *
	 * @var array
	 */
	private $_isSpannedCell	= array();

	/**
	 * Excel cells that are upper-left corner in a cell merge
	 *
	 * @var array
	 */
	private $_isBaseCell	= array();

	/**
	 * Excel rows that should not be written as HTML rows
	 *
	 * @var array
	 */
	private $_isSpannedRow	= array();

	/**
	 * Is the current writer creating PDF?
	 *
	 * @var boolean
	 */
	protected $_isPdf = true;

	/**
	 * Generate the Navigation block
	 *
	 * @var boolean
	 */
	private $_generateSheetNavigationBlock = true;
    /**
     * @var Pager
     */
    protected $pager ;
    protected $isSplitTables = false;
    protected $isUseCellWidth = false;

    /**
     * ��������� �� ����� �� �������� �������
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
     * ������������ td.width
     * @param $isUseCellWidth
     * @return $this
     */
    protected function setIsUseCellWidth($isUseCellWidth)
    {
        $this->isUseCellWidth = $isUseCellWidth;
        return $this;
    }

    protected function getIsUseCellWidth()
    {
        return $this->isUseCellWidth;
    }

    public function setPager(Pager $pager=null)
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
            $this->pager = new DefaultPager($this->_phpExcel->getActiveSheet());
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
     * ������ ��������� � �������
     * @param $sheet
     * @return int
     */
    private function getDocumentWidth($sheetIndex){

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
            if ($this->getIsUseCellWidth()){
                $style = 'style="overflow: wrap; "';
            }
            else {
                $style = 'style="width:'.$this->getDocumentWidth($sheetIndex).'pt "';
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
    private function buildRowHtml($sheet, $row,  $dimension){
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
     * ����������� �������
     * @return string
     */
    private function getPageBreakHtml(){
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

                        $rowHtml = $this->buildRowHtml($sheet, $row, $dimension);
                        $html .= $rowHtml;
                        if ($row>=$theadStart && $row<=$theadEnd){
                            $headerHtml .=$rowHtml ;
                        }
                        if ($row == $theadEnd) {

                            if (true == $this->getIsSplitTables()){
                                $this->_generateTableFooter();
                                $this->_generateTableHeader($sheet);
                            }
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
