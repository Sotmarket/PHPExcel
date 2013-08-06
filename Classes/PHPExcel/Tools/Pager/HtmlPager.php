<?php
/**
 * Пейджер для html.
 * @category  
 * @package   
 * @subpackage 
 * @author: u.lebedev
 * @date: 06.08.13
 * @version    $Id: $
 */
class PHPExcel_Tools_Pager_HtmlPager  extends PHPExcel_Tools_Pager_Pager implements IExcelPager{

    const PRECISION_ROW=5;

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
                if ($sum > (($this->getPageHeight()-80))){

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

}