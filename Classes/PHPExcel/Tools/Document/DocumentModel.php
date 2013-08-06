<?php
/**
 * 
 * @category  
 * @package   
 * @subpackage 
 * @author: u.lebedev
 * @date: 30.07.13
 * @version    $Id: $
 */

class PHPExcel_Tools_Document_DocumentModel
{
    /**
     * Left
     *
     * @var double
     */
    private $left;

    /**
     * Right
     *
     * @var double
     */
    private $right;

    /**
     * Top
     *
     * @var double
     */
    private $top;

    /**
     * Bottom
     *
     * @var double
     */
    private $bottom;

    /**
     * Header
     *
     * @var double
     */
    private $header;

    /**
     * Footer
     *
     * @var double
     */
    private $footer ;
    protected $format;
    protected $orientation=self::ORIENTATION_PORTRAIT;
    const ORIENTATION_PORTRAIT = "P";
    const ORIENTATION_LANDSCAPE = "L";
    const INCH_FACTOR = 25.4; // inches to mm
    /**
     * Create a new PageModel
     */
    public function __construct(
        $format,
        $orientation,
        $left=null,
        $right=null,
        $top=null,
        $bottom=null,
        $header=null,
        $footer=null

    )
    {
        $this
            ->setFormat($format)
            ->setOrientation($format)
            ->setTop($top)
            ->setBottom($bottom)
            ->setLeft($left)
            ->setRight($right)
            ->setHeader($header)
            ->setFooter($footer)
        ;
    }

    /**
     * @param PHPExcel_Worksheet $sheet
     * @return DocuemntModel
     */
    public static function getByExcelSheet(PHPExcel_Worksheet $sheet){

        $pageSetup =$sheet->getPageSetup();
        $orientation = ($pageSetup->getOrientation() == PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE)
            ? self::ORIENTATION_LANDSCAPE : self::ORIENTATION_PORTRAIT;
        $printPaperSize = $pageSetup->getPaperSize();
        if (isset(PHPExcel_Writer_PDF_Core::$_paperSizes[$printPaperSize])){
            $paperType = PHPExcel_Writer_PDF_Core::$_paperSizes[$printPaperSize];
        }
        else {
            throw new Exception ("Unknown paperSize" . $printPaperSize);
        }
        $printMargins = $sheet->getPageMargins();
        $instance = new self (
            $paperType,
            $orientation,
            $printMargins->getLeft()*self::INCH_FACTOR, // milimetres
            $printMargins->getRight()*self::INCH_FACTOR,
            $printMargins->getTop()*self::INCH_FACTOR,
            $printMargins->getBottom()*self::INCH_FACTOR,
            $printMargins->getHeader()*self::INCH_FACTOR,
            $printMargins->getFooter()*self::INCH_FACTOR
        );
        return $instance;
    }

    /**
     * Получить высоту документа
     * @return int
     * @throws Exception
     */
    public function getHeight(){
        $arrHeightWidth = $this->getPageFormat();
        if (false === $arrHeightWidth){
            throw new Exception ("Unknown format:" . $this->getFormat());
        }
        $result = 0;
        if ($this->getOrientation() == self::ORIENTATION_PORTRAIT){
            $result = $arrHeightWidth[1];
        }
        else {
            $result = $arrHeightWidth[0];
        }

        return $result;

    }

    /**
     * Получить длину документа
     * @return int
     * @throws Exception
     */
    public function getWidth(){
        $arrHeightWidth = $this->getPageFormat();
        if (false === $arrHeightWidth){
            throw new Exception ("Unknown format:" . $this->getFormat());
        }
        $result = 0;
        if ($this->getOrientation() == self::ORIENTATION_PORTRAIT){
            $result = $arrHeightWidth[0];
        }
        else {
            $result = $arrHeightWidth[1];
        }

        return $result;
    }
    /**
     * @param $format
     * @return PHPExcel_Tools_Document_DocumentModel
     */
    protected function setFormat($format)
    {
        $this->format = $format;
        return $this;
    }

    /**
     * @param $orientation
     * @return PHPExcel_Tools_Document_DocumentModel
     */
    protected function setOrientation($orientation)
    {
        $this->orientation = $orientation;
        return $this;
    }

    protected function getOrientation()
    {
        return $this->orientation;
    }

    protected function getFormat()
    {
        return $this->format;
    }

    /**
     * Получить размеры документа
     * @return array|bool
     */
    function getPageFormat() {
        $format = $this->getFormat();
        switch (strtoupper($format)) {
            case '4A0': {$format = array(4767.87,6740.79); break;}
            case '2A0': {$format = array(3370.39,4767.87); break;}
            case 'A0': {$format = array(2383.94,3370.39); break;}
            case 'A1': {$format = array(1683.78,2383.94); break;}
            case 'A2': {$format = array(1190.55,1683.78); break;}
            case 'A3': {$format = array(841.89,1190.55); break;}
            case 'A4': default: {$format = array(595.28,841.89); break;}
            case 'A5': {$format = array(419.53,595.28); break;}
            case 'A6': {$format = array(297.64,419.53); break;}
            case 'A7': {$format = array(209.76,297.64); break;}
            case 'A8': {$format = array(147.40,209.76); break;}
            case 'A9': {$format = array(104.88,147.40); break;}
            case 'A10': {$format = array(73.70,104.88); break;}
            case 'B0': {$format = array(2834.65,4008.19); break;}
            case 'B1': {$format = array(2004.09,2834.65); break;}
            case 'B2': {$format = array(1417.32,2004.09); break;}
            case 'B3': {$format = array(1000.63,1417.32); break;}
            case 'B4': {$format = array(708.66,1000.63); break;}
            case 'B5': {$format = array(498.90,708.66); break;}
            case 'B6': {$format = array(354.33,498.90); break;}
            case 'B7': {$format = array(249.45,354.33); break;}
            case 'B8': {$format = array(175.75,249.45); break;}
            case 'B9': {$format = array(124.72,175.75); break;}
            case 'B10': {$format = array(87.87,124.72); break;}
            case 'C0': {$format = array(2599.37,3676.54); break;}
            case 'C1': {$format = array(1836.85,2599.37); break;}
            case 'C2': {$format = array(1298.27,1836.85); break;}
            case 'C3': {$format = array(918.43,1298.27); break;}
            case 'C4': {$format = array(649.13,918.43); break;}
            case 'C5': {$format = array(459.21,649.13); break;}
            case 'C6': {$format = array(323.15,459.21); break;}
            case 'C7': {$format = array(229.61,323.15); break;}
            case 'C8': {$format = array(161.57,229.61); break;}
            case 'C9': {$format = array(113.39,161.57); break;}
            case 'C10': {$format = array(79.37,113.39); break;}
            case 'RA0': {$format = array(2437.80,3458.27); break;}
            case 'RA1': {$format = array(1729.13,2437.80); break;}
            case 'RA2': {$format = array(1218.90,1729.13); break;}
            case 'RA3': {$format = array(864.57,1218.90); break;}
            case 'RA4': {$format = array(609.45,864.57); break;}
            case 'SRA0': {$format = array(2551.18,3628.35); break;}
            case 'SRA1': {$format = array(1814.17,2551.18); break;}
            case 'SRA2': {$format = array(1275.59,1814.17); break;}
            case 'SRA3': {$format = array(907.09,1275.59); break;}
            case 'SRA4': {$format = array(637.80,907.09); break;}
            case 'LETTER': {$format = array(612.00,792.00); break;}
            case 'LEGAL': {$format = array(612.00,1008.00); break;}
            case 'LEDGER': {$format = array(279.00,432.00); break;}
            case 'TABLOID': {$format = array(279.00,432.00); break;}
            case 'EXECUTIVE': {$format = array(521.86,756.00); break;}
            case 'FOLIO': {$format = array(612.00,936.00); break;}
            case 'B': {$format=array(362.83,561.26 );	 break;}		//	'B' format paperback size 128x198mm
            case 'A': {$format=array(314.65,504.57 );	 break;}		//	'A' format paperback size 111x178mm
            case 'DEMY': {$format=array(382.68,612.28 );  break;}		//	'Demy' format paperback size 135x216mm
            case 'ROYAL': {$format=array(433.70,663.30 );  break;}	//	'Royal' format paperback size 153x234mm
            default: $format = false;
        }
        return $format;
    }

    /**
     * Get Left
     *
     * @return double
     */
    public function getLeft() {
        return $this->left;
    }

    /**
     * Set Left
     *
     * @param double $pValue
     * @return PHPExcel_Tools_Document_DocumentModel
     */
    public function setLeft($pValue) {
        $this->left = $pValue;
        return $this;
    }

    /**
     * Get Right
     *
     * @return double
     */
    public function getRight() {
        return $this->right;
    }

    /**
     * Set Right
     *
     * @param double $pValue
     * @return PHPExcel_Tools_Document_DocumentModel
     */
    public function setRight($pValue) {
        $this->right = $pValue;
        return $this;
    }

    /**
     * Get Top
     *
     * @return double
     */
    public function getTop() {
        return $this->top;
    }

    /**
     * Set Top
     *
     * @param double $pValue
     * @return PageModel
     */
    public function setTop($pValue) {
        $this->top = $pValue;
        return $this;
    }

    /**
     * Get Bottom
     *
     * @return double
     */
    public function getBottom() {
        return $this->bottom;
    }

    /**
     * Set Bottom
     *
     * @param double $pValue
     * @return PHPExcel_Tools_Document_DocumentModel
     */
    public function setBottom($pValue) {
        $this->bottom = $pValue;
        return $this;
    }

    /**
     * Get Header
     *
     * @return double
     */
    public function getHeader() {
        return $this->header;
    }

    /**
     * Set Header
     *
     * @param double $pValue
     * @return PageModel
     */
    public function setHeader($pValue) {
        $this->header = $pValue;
        return $this;
    }

    /**
     * Get Footer
     *
     * @return double
     */
    public function getFooter() {
        return $this->footer;
    }

    /**
     * Set Footer
     *
     * @param double $pValue
     * @return PHPExcel_Tools_Document_DocumentModel
     */
    public function setFooter($pValue) {
        $this->footer = $pValue;
        return $this;
    }

    /**
     * Implement PHP __clone to create a deep clone, not just a shallow copy.
     */
    public function __clone() {
        $vars = get_object_vars($this);
        foreach ($vars as $key => $value) {
            if (is_object($value)) {
                $this->$key = clone $value;
            } else {
                $this->$key = $value;
            }
        }
    }
}
