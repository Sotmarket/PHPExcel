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
    protected $isUseWordWrap=true;
    protected $widthErrorFactor = 1;

    /**
     * Определение ширины таблицы происходит с ошибкой, видимо разный алгоритм интерпретации размера шраифтов
     * @param $widthErrorFactor
     * @return $this
     */
    public function setWidthErrorFactor($widthErrorFactor)
    {
        $this->widthErrorFactor = $widthErrorFactor;
        return $this;
    }

    protected function getWidthErrorFactor()
    {
        return $this->widthErrorFactor;
    }

    protected function setIsUseWordWrap($isUseWordWrap)
    {
        $this->isUseWordWrap = $isUseWordWrap;
        return $this;
    }

    protected function getIsUseWordWrap()
    {
        return $this->isUseWordWrap;
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
            if ($this->getIsUseWordWrap()){
               // $width = $this->getDocumentWidth($sheetIndex)*$this->getWidthErrorFactor();
                $style = 'style="overflow: wrap; table-layout: fixed;width:100%"';
            }
            else {
                $style="";
            }

            $html .= '	<table border="0" cellpadding="0" cellspacing="0" '.$style.' id="sheet' . $sheetIndex . '" class="sheet' . $sheetIndex . $gridlines . '">' . PHP_EOL;
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

    /**
     * Generate row
     *
     * @param	PHPExcel_Worksheet	$pSheet			PHPExcel_Worksheet
     * @param	array				$pValues		Array containing cells in a row
     * @param	int					$pRow			Row number (0-based)
     * @return	string
     * @throws	PHPExcel_Writer_Exception
     */
    protected function _generateRow(PHPExcel_Worksheet $pSheet, $pValues = null, $pRow = 0) {
        if (is_array($pValues)) {
            // Construct HTML
            $html = '';

            // Sheet index
            $sheetIndex = $pSheet->getParent()->getIndex($pSheet);

            // DomPDF and breaks
            if ($this->_isPdf && count($pSheet->getBreaks()) > 0) {
                $breaks = $pSheet->getBreaks();

                // check if a break is needed before this row
                if (isset($breaks['A' . $pRow])) {
                    // close table: </table>
                    $html .= $this->_generateTableFooter();

                    // insert page break
                    $html .= '<div style="page-break-before:always" />';

                    // open table again: <table> + <col> etc.
                    $html .= $this->_generateTableHeader($pSheet);
                }
            }

            // Write row start

            if (!$this->_useInlineCss) {
                $html .= '		  <tr class="row' . $pRow . '" >' . PHP_EOL;
            } else {
                $style = isset($this->_cssStyles['table.sheet' . $sheetIndex . ' tr.row' . $pRow])
                    ? $this->_assembleCSS($this->_cssStyles['table.sheet' . $sheetIndex . ' tr.row' . $pRow]) : '';

                $html .= '		  <tr style="' . $style . '">' . PHP_EOL;
            }

            // Write cells
            $colNum = 0;
            foreach ($pValues as $cell) {
                $coordinate = PHPExcel_Cell::stringFromColumnIndex($colNum) . ($pRow + 1);

                if (!$this->_useInlineCss) {
                    $cssClass = '';
                    $cssClass = 'col' . $colNum;
                } else {
                    $cssClass = array();
                    if (isset($this->_cssStyles['table.sheet' . $sheetIndex . ' td.column' . $colNum])) {
                        $this->_cssStyles['table.sheet' . $sheetIndex . ' td.column' . $colNum];
                    }
                }
                $colSpan = 1;
                $rowSpan = 1;

                // initialize
                $cellData = '&nbsp;';

                // PHPExcel_Cell
                if ($cell instanceof PHPExcel_Cell) {
                    $cellData = '';
                    if (is_null($cell->getParent())) {
                        $cell->attach($pSheet);
                    }
                    // Value
                    if ($cell->getValue() instanceof PHPExcel_RichText) {
                        // Loop through rich text elements
                        $elements = $cell->getValue()->getRichTextElements();
                        foreach ($elements as $element) {
                            // Rich text start?
                            if ($element instanceof PHPExcel_RichText_Run) {
                                $cellData .= '<span style="' . $this->_assembleCSS($this->_createCSSStyleFont($element->getFont())) . '">';

                                if ($element->getFont()->getSuperScript()) {
                                    $cellData .= '<sup>';
                                } else if ($element->getFont()->getSubScript()) {
                                    $cellData .= '<sub>';
                                }
                            }

                            // Convert UTF8 data to PCDATA
                            $cellText = $element->getText();
                            $cellData .= htmlspecialchars($cellText);

                            if ($element instanceof PHPExcel_RichText_Run) {
                                if ($element->getFont()->getSuperScript()) {
                                    $cellData .= '</sup>';
                                } else if ($element->getFont()->getSubScript()) {
                                    $cellData .= '</sub>';
                                }

                                $cellData .= '</span>';
                            }
                        }
                    } else {
                        if ($this->_preCalculateFormulas) {
                            $cellData = PHPExcel_Style_NumberFormat::toFormattedString(
                                $cell->getCalculatedValue(),
                                $pSheet->getParent()->getCellXfByIndex( $cell->getXfIndex() )->getNumberFormat()->getFormatCode(),
                                array($this, 'formatColor')
                            );
                        } else {
                            $cellData = PHPExcel_Style_NumberFormat::ToFormattedString(
                                $cell->getValue(),
                                $pSheet->getParent()->getCellXfByIndex( $cell->getXfIndex() )->getNumberFormat()->getFormatCode(),
                                array($this, 'formatColor')
                            );
                        }
                        $cellData = htmlspecialchars($cellData);
                        if ($pSheet->getParent()->getCellXfByIndex( $cell->getXfIndex() )->getFont()->getSuperScript()) {
                            $cellData = '<sup>'.$cellData.'</sup>';
                        } elseif ($pSheet->getParent()->getCellXfByIndex( $cell->getXfIndex() )->getFont()->getSubScript()) {
                            $cellData = '<sub>'.$cellData.'</sub>';
                        }
                    }

                    // Converts the cell content so that spaces occuring at beginning of each new line are replaced by &nbsp;
                    // Example: "  Hello\n to the world" is converted to "&nbsp;&nbsp;Hello\n&nbsp;to the world"
                    $cellData = preg_replace("/(?m)(?:^|\\G) /", '&nbsp;', $cellData);

                    // convert newline "\n" to '<br>'
                    $cellData = nl2br($cellData);

                    // Extend CSS class?
                    if (!$this->_useInlineCss) {
                        $cssClass .= ' style' . $cell->getXfIndex();
                        $cssClass .= ' ' . $cell->getDataType();
                    } else {
                        if (isset($this->_cssStyles['td.style' . $cell->getXfIndex()])) {
                            $cssClass = array_merge($cssClass, $this->_cssStyles['td.style' . $cell->getXfIndex()]);
                        }

                        // General horizontal alignment: Actual horizontal alignment depends on dataType
                        $sharedStyle = $pSheet->getParent()->getCellXfByIndex( $cell->getXfIndex() );
                        if ($sharedStyle->getAlignment()->getHorizontal() == PHPExcel_Style_Alignment::HORIZONTAL_GENERAL
                            && isset($this->_cssStyles['.' . $cell->getDataType()]['text-align']))
                        {
                            $cssClass['text-align'] = $this->_cssStyles['.' . $cell->getDataType()]['text-align'];
                        }
                    }
                }

                // Hyperlink?
                if ($pSheet->hyperlinkExists($coordinate) && !$pSheet->getHyperlink($coordinate)->isInternal()) {
                    $cellData = '<a href="' . htmlspecialchars($pSheet->getHyperlink($coordinate)->getUrl()) . '" title="' . htmlspecialchars($pSheet->getHyperlink($coordinate)->getTooltip()) . '">' . $cellData . '</a>';
                }

                // Should the cell be written or is it swallowed by a rowspan or colspan?
                $writeCell = ! ( isset($this->_isSpannedCell[$pSheet->getParent()->getIndex($pSheet)][$pRow + 1][$colNum])
                    && $this->_isSpannedCell[$pSheet->getParent()->getIndex($pSheet)][$pRow + 1][$colNum] );

                // Colspan and Rowspan
                $colspan = 1;
                $rowspan = 1;


                if (isset($this->_isBaseCell[$pSheet->getParent()->getIndex($pSheet)][$pRow + 1][$colNum])) {

                    $spans = $this->_isBaseCell[$pSheet->getParent()->getIndex($pSheet)][$pRow + 1][$colNum];
                    $rowSpan = $spans['rowspan'];
                    $colSpan = $spans['colspan'];
                    //	Also apply style from last cell in merge to fix borders -
                    //		relies on !important for non-none border declarations in _createCSSStyleBorder
                    $endCellCoord = PHPExcel_Cell::stringFromColumnIndex($colNum + $colSpan - 1) . ($pRow + $rowSpan);

                    if (is_array($cssClass)){
                        $cssClass["class"] = ' style' . $pSheet->getCell($endCellCoord)->getXfIndex();
                    }
                    else {
                        $cssClass .= ' style' . $pSheet->getCell($endCellCoord)->getXfIndex();
                    }


                }

                // Write
                if ($writeCell) {
                    // Column start
                    $html .= '			<td';
                    if (!$this->_useInlineCss) {
                        $style="style='word-wrap: break-word'";
                        $html .= ' class="' . $cssClass . '" style';
                    } else {
                        //** Necessary redundant code for the sake of PHPExcel_Writer_PDF **
                        // We must explicitly write the width of the <td> element because TCPDF
                        // does not recognize e.g. <col style="width:42pt">
                        $width = 0;
                        $i = $colNum - 1;
                        $e = $colNum + $colSpan - 1;
                        while($i++ < $e) {
                            if (isset($this->_columnWidths[$sheetIndex][$i])) {
                                $width += $this->_columnWidths[$sheetIndex][$i];
                            }
                        }
                        if (is_array($cssClass)){
                            $cssClass['width'] = $width . 'pt';
                        }


                        // We must also explicitly write the height of the <td> element because TCPDF
                        // does not recognize e.g. <tr style="height:50pt">
                        if (isset($this->_cssStyles['table.sheet' . $sheetIndex . ' tr.row' . $pRow]['height'])) {
                            $height = $this->_cssStyles['table.sheet' . $sheetIndex . ' tr.row' . $pRow]['height'];
                            if (is_array($cssClass)){
                                $cssClass['height'] = $height;
                            }

                        }

                        //** end of redundant code **
                        if (isset($cssClass["class"])){
                            $html .= ' class="' . trim($cssClass["class"]) . '"';
                            unset($cssClass["class"]);
                        }
                        $html .= ' style="' . $this->_assembleCSS($cssClass) . ' "';
                    }
                    if ($colSpan > 1) {
                        $html .= ' colspan="' . $colSpan . '"';
                    }
                    if ($rowSpan > 1) {
                        $html .= ' rowspan="' . $rowSpan . '"';
                    }
                    $html .= '>';

                    // Image?
                    $html .= $this->_writeImageInCell($pSheet, $coordinate);

                    // Chart?
                    if ($this->_includeCharts) {
                        $html .= $this->_writeChartInCell($pSheet, $coordinate);
                    }

                    // Cell data
                    $html .= $cellData;

                    // Column end
                    $html .= '</td>' . PHP_EOL;
                }

                // Next column
                ++$colNum;
            }

            // Write row end
            $html .= '		  </tr>' . PHP_EOL;

            // Return
            return $html;
        } else {
            throw new PHPExcel_Writer_Exception("Invalid parameters passed.");
        }
    }

}
