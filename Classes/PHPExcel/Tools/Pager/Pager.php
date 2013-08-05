<?php
/**
 *
 * @category
 * @package
 * @subpackage
 * @author: u.lebedev
 * @date: 26.07.13
 * @version    $Id: $
 */
class Pager implements IExcelPager{
    protected $excelSheet;

    /**
     * ������ ��������
     * @var integer
     */
    protected $pageHeight;
    /**
     * ��������� �������� �� ���������
     * @var array
     */
    protected $rowBounds = array();
    protected $lastDataRow;
    const INCH_FACTOR = 25.4; // inches to mm
    // ���������� ������� - ������� �� �������� ����� � ���������
    const PRECISION_ROW=4;

    /**
     *
     * @param PHPExcel_Worksheet $excelSheet
     * @param null               $lastDataRow ��������� ��� � �������
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
     * ���������� ������ ��������
     * @return int
     */
    public function getPageHeight()
    {
        if (null == $this->pageHeight){
            $pageModel = DocumentModel::getByExcelSheet($this->getExcelSheet());
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
     * �������� ��������������� ��������
     * @return array
     */
    public function getRowBounds()
    {
        if (0 == count ($this->rowBounds)){
            $sheet = $this->getExcelSheet();
            $highestRow = $sheet->getHighestRow()+2; // ����� � �������
            $page = 1;
            $printMargins = $sheet->getPageMargins();
            $footer_top  =($printMargins->getTop()+$printMargins->getBottom())*self::INCH_FACTOR;

            $sum = $footer_top;
            // $footer_top  =25;
            // ����� ����� ������� ���������, ������������ �� ������ ��������
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

    /**
     * �������� ������ �������
     * @return array
     */
    public function getSmoothedPageMap(){
        $rowBounds = $this->getRowBounds();
        if (null  == $this->getLastDataRow()){
            $result = array(1=>new PageModel(1, $rowBounds[1]));
            return $result;
        }

        $countPages = $this->getCountPages();
        $printMargins = $this->getExcelSheet()->getPageMargins();
        $lastDataRow = $this->getLastDataRow();
        $footer_top = ($printMargins->getTop()+$printMargins->getBottom())*self::INCH_FACTOR;
        if (1 == $countPages){
            $result = array(1=>new PageModel(1, $rowBounds[1]));
        }else {
            $result = array();
            $transfer = self::PRECISION_ROW;
            $cutDown = false;
            foreach ($rowBounds as $page=>$finishRow){
                if (1 == $page){
                    $startRow = 1;
                }
                else {
                    $startRow = $result[$page-1]->getFinish()+1;
                }

                if ($page < $countPages){
                    $pageModel = new PageModel($startRow, $finishRow-$transfer);
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

                            $pageModel->finish = $lastDataRow-1-self::PRECISION_ROW;
                            $cutDown = true;

                        }

                    }
                    $result[$page]=$pageModel;
                }
                else {
                    // ��������� ��������
                    $pageModel = new PageModel($startRow, $finishRow-$transfer);
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
                            $pageModel->finish = $lastDataRow-1-self::PRECISION_ROW;
                            $lastPage = new PageModel($lastDataRow-self::PRECISION_ROW, $finishRow);
                            $result[$page+1]=$lastPage;
                        }

                    }
                    $result[$page] = $pageModel;

                }
            }
        }
        ksort($result);
        $maxRow = $this->getExcelSheet()->getHighestRow();
        for($i=1; $i<=count($result); $i++){
            if ($result[$i]->getFinish()>$maxRow){
                $result[$i]->finish = $maxRow;
            }
            if ($result[$i]->getStart()>$maxRow){
                unset($result[$i]);
            }

        }
        return $result;
    }

    /**
     * �������� ������ ����������
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
     * �������� ������ ��������� �����
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
     * �������� ������ ����
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
     * ���� �� ������ ������ ���� � ��������� - �������� �� ���������
     * @return float
     */
    public function getDefaultRowHeight (){

        return PHPExcel_Shared_Font::getDefaultRowHeightByFont($this->getExcelSheet()->getDefaultStyle()->getFont());
    }

    /**
     * ��������� ������ ���������
     * @return double|float|int
     */
    public function getFullDocumentHeight(){
        return $this->getHeightOfRowRange(1, null);
    }

    /**
     * �������� ���������� �������
     * @return int
     */
    public function getCountPages(){
        return count($this->getRowBounds());
    }

    /**
     * �� ����� �������� ��������� ���
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