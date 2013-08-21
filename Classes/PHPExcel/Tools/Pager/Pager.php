<?php
require_once ("IExcelPager.php");
/**
 *
 * @category
 * @package
 * @subpackage
 * @author: u.lebedev
 * @date: 26.07.13
 * @version    $Id: $
 */
class PHPExcel_Tools_Pager_Pager implements IExcelPager{
    protected $excelSheet;

    /**
     * Высота страницы
     * @var integer
     */
    protected $pageHeight;
    /**
     * Примерная разбивка по страницам
     * @var array
     */
    protected $rowBounds = array();
    protected $lastDataRow;
    const INCH_FACTOR = 25.4; // inches to mm
    // колчиество колонок - которые мы отрезаем снизу у документа
    const PRECISION_ROW=4;
    protected $smoothedPageMap = array();

    /**
     *
     * @param PHPExcel_Worksheet $excelSheet
     * @param null               $lastDataRow Послдений ряд с данными
     */
    public function __construct(PHPExcel_Worksheet $excelSheet, $lastDataRow=null){
        $this
            ->setExcelSheet($excelSheet)
            ->setLastDataRow($lastDataRow)
        ;
    }

    /**
     * @param integer $lastDataRow
     * @return $this
     */
    protected function setLastDataRow(  $lastDataRow)
    {
        $this->lastDataRow = $lastDataRow;
        return $this;
    }

    /**
     * @return integer
     */
    protected function getLastDataRow()
    {
        return $this->lastDataRow;
    }

    /**
     * @param $pageHeight
     * @return $this
     */
    protected function setPageHeight($pageHeight)
    {
        $this->pageHeight = $pageHeight;
        return $this;
    }

    /**
     * Вычисление высоты страницы
     * @return int
     */
    public function getPageHeight()
    {
        if (null == $this->pageHeight){
            $pageModel = PHPExcel_Tools_Document_DocumentModel::getByExcelSheet($this->getExcelSheet());
            $this->pageHeight = $pageModel->getHeight();
        }
        return $this->pageHeight;
    }

    /**
     * @param $rowBounds
     * @return $this
     */
    public function setRowBounds($rowBounds)
    {
        $this->rowBounds = $rowBounds;
        return $this;
    }

    /**
     * Получить предварительную разбивку
     * @return array
     */
    public function getRowBounds()
    {
        if (0 == count ($this->rowBounds)){
            $sheet = $this->getExcelSheet();
            $highestRow = $sheet->getHighestRow()+2; // футер с запасом
            $page = 1;
            $printMargins = $sheet->getPageMargins();
            $footer_top  =($printMargins->getTop()+$printMargins->getBottom())*self::INCH_FACTOR;

            $sum = $footer_top;
            // $footer_top  =25;
            // сумма высот колонок заголовка, переносимого на каждой странице
            $header = $this->getHeaderRowHeight();

            $count_per_page = 0;
            for($i=1; $i<=$highestRow; $i++){
                $height = $this->getRowHeight($i);

                $sum+=$height;
                $count_per_page++;
                if ($sum > (($this->getPageHeight()))){

                    $this->rowBounds[$page]=$i-1;
                    $page++;
                    $sum=$height+$footer_top+$header;
                    $count_per_page = 0;

                }

            }
            // last page

            $this->rowBounds[$page]=$i;


        }
        return $this->rowBounds;
    }
    private function getPrecisionByPage($page){
        $result = 1;
        if (1 < $page){
            $result = self::PRECISION_ROW;
        }
        return $result;
    }

    /**
     * Установить карту страничек
     * @param $pageMap
     * @return $this
     */
    public function setSmoothedPageMap($pageMap){
        $this->smoothedPageMap = $pageMap;
        return $this;
    }
    /**
     * Получить список страниц
     * @return array
     */
    public function getSmoothedPageMap(){
        if (null == $this->smoothedPageMap){
            $rowBounds = $this->getRowBounds();
            if (null  == $this->getLastDataRow()){
                $result = array(1=>new PHPExcel_Tools_Document_PageModel(1, $rowBounds[1]));
                return $result;
            }

            $countPages = $this->getCountPages();
            $printMargins = $this->getExcelSheet()->getPageMargins();
            $lastDataRow = $this->getLastDataRow();
            $footer_top = ($printMargins->getTop()+$printMargins->getBottom())*self::INCH_FACTOR;
            if (1 == $countPages){
                $result = array(1=>new PHPExcel_Tools_Document_PageModel(1, $rowBounds[1]));
            }else {
                $result = array();

                $cutDown = false;
                foreach ($rowBounds as $page=>$finishRow){
                    $transfer = $this->getPrecisionByPage($page);
                    if (1 == $page){
                        $startRow = 1;
                    }
                    else {
                        $startRow = $result[$page-1]->getFinish()+1;
                    }

                    if ($page < $countPages){
                        $pageModel = new PHPExcel_Tools_Document_PageModel($startRow, $finishRow-$transfer);
                        if (
                            $lastDataRow>$pageModel->getStart() &&
                            ($lastDataRow<=$pageModel->getFinish() || $lastDataRow==($pageModel->getFinish()+1))
                        ){
                            $headerHeight = ($page>1)?$this->getHeaderRowHeight():0;
                            $haveHeight = $this->getPageHeight()
                                - $this->getHeightOfRowRange($pageModel->getStart(), $lastDataRow)
                                - $footer_top
                                - $headerHeight
                            ;
                            $needHeight = $this->getHeightOfRowRange($lastDataRow);

                            if ($needHeight>$haveHeight){

                                $pageModel->finish = $lastDataRow-1-$transfer;
                                $cutDown = true;

                            }

                        }
                        $result[$page]=$pageModel;
                    }
                    else {
                        // Последняя страница
                        $pageModel = new PHPExcel_Tools_Document_PageModel($startRow, $finishRow-$transfer);
                        if (
                            $lastDataRow>$pageModel->getStart() &&
                            ($lastDataRow<=$pageModel->getFinish() || $lastDataRow==($pageModel->getFinish()+1)) &&
                            !$cutDown
                        )
                        {
                            $headerHeight = ($page>1)?$this->getHeaderRowHeight():0;
                            $haveHeight = $this->getPageHeight()
                                - $this->getHeightOfRowRange($pageModel->getStart(), $lastDataRow)
                                - $footer_top
                                - $headerHeight
                            ;
                            $needHeight = $this->getHeightOfRowRange($lastDataRow);

                            if ($needHeight>$haveHeight){
                                //$transfer = $pageModel->finish -($lastDataRow-1);
                                $pageModel->finish = $lastDataRow-1-$transfer;
                                $lastPage = new PHPExcel_Tools_Document_PageModel(
                                    $lastDataRow-$transfer, $finishRow
                                );
                                $result[$page+1]=$lastPage;
                            }

                        }
                        $result[$page] = $pageModel;

                    }
                }
            }
            ksort($result);
            $maxRow = $this->getExcelSheet()->getHighestRow();
           // print_r($result); die();
            for($i=1; $i<=count($result); $i++){
                if ($result[$i]->getFinish()>$maxRow){
                    $result[$i]->finish = $maxRow;
                }

                if ($result[$i]->getFinish() <= $result[$i]->getStart() ){
                    if (isset($result[$i+1])){
                        $result[$i]->finish = $result[$i]->start+1;
                        $result[$i+1]->start = $result[$i]->start+2;
                    }
                    else {
                        unset($result[$i]);
                    }

                }
                if (isset($result[$i]) && $result[$i]->getStart()>$maxRow){
                    unset($result[$i]);
                }

            }

            $this->smoothedPageMap = $result;
        }

        return $this->smoothedPageMap;
    }

    /**
     * Получить размер заголовков
     * @return double|int
     */
    public function getHeaderRowHeight(){
        $sheet = $this->getExcelSheet();
        $rowsRepeat = $sheet->getPageSetup()->getRowsToRepeatAtTop();
        $result = 0;
        if (is_array($rowsRepeat) && $rowsRepeat[0]>0){

            for ($k=$rowsRepeat[0]; $k<=$rowsRepeat[1]; $k++){
                $result+= $sheet->getRowDimension($k)->getRowHeight();
            }


        }
        // $result+=60;
        return $result;
    }

    /**
     * Получить высоты диапазона рядов
     * @param int  $rowStart
     * @param null $rowFinish
     * @return double|float|int
     */
    public function getHeightOfRowRange($rowStart =1, $rowFinish=null){
        $sheet = $this->getExcelSheet();
        if (null == $rowFinish){
            $highestRow = $sheet->getHighestRow();
        }
        else {
            $highestRow = $rowFinish;
        }
        $sum = 0;
        for($i=$rowStart; $i<=$highestRow; $i++){
            $height = $this->getRowHeight($i);
            $sum+= $height;
        }
        return $sum;
    }

    /**
     * Получить высоту ряда
     * @param $row
     *
     * @return double|float
     */
    protected function getRowHeight ($row){
        //$row = $row-1;
        $sheet = $this->getExcelSheet();
        $dimensions = $sheet->getRowDimensions();

        if (isset($dimensions[$row])){
            $height = $dimensions[$row]->getRowHeight( );
        }
        else{
            $height = $this->getDefaultRowHeight();
        }
        return $height;
    }

    /**
     * Если не задана высота ряда в документе - получить по умолчанию
     * @return float
     */
    public function getDefaultRowHeight (){

        return PHPExcel_Shared_Font::getDefaultRowHeightByFont($this->getExcelSheet()->getDefaultStyle()->getFont());
    }

    /**
     * Посчитать высоту документа
     * @return double|float|int
     */
    public function getFullDocumentHeight(){
        return $this->getHeightOfRowRange(1, null);
    }

    /**
     * Получить количество страниц
     * @return int
     */
    public function getCountPages(){
        return count($this->getRowBounds());
    }

    /**
     * На какой странице находится ряд
     * @param $row
     * @return int|string
     */
    public function getPageOfRow($row){
        //$row = $row-1;
        $rowBounds = $this->getRowBounds();
        $result = 1;
        foreach ($rowBounds as $page=>$finishRow){
            if ($row<=$finishRow){
                $result = $page;
                break;
            }
        }
        return $result;
    }

    /**
     * @param PHPExcel_Worksheet $excelSheet
     * @return $this
     */
    protected function setExcelSheet( PHPExcel_Worksheet $excelSheet)
    {
        $this->excelSheet = $excelSheet;
        return $this;
    }

    /**
     * @return PHPExcel_Worksheet
     */
    protected function getExcelSheet()
    {
        return $this->excelSheet;
    }

}