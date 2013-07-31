<?php
/**
 * PHPExcel
 *
 * Copyright (c) 2006 - 2012 PHPExcel
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 *
 * @category	PHPExcel
 * @package		PHPExcel_Writer
 * @copyright	Copyright (c) 2006 - 2012 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license		http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version		1.7.8, 2012-10-12
 */


/** Require mPDF library */
$pdfRendererClassFile = PHPExcel_Settings::getPdfRendererPath() . '/mpdf.php';
if (file_exists($pdfRendererClassFile)) {
	require_once $pdfRendererClassFile;
} else {
	throw new Exception('Unable to load PDF Rendering library');
}

/**
 * PHPExcel_Writer_PDF_mPDF
 *
 * @category	PHPExcel
 * @package		PHPExcel_Writer
 * @copyright	Copyright (c) 2006 - 2012 PHPExcel (http://www.codeplex.com/PHPExcel)
 */
class PHPExcel_Writer_PDF_mPDF extends PHPExcel_Writer_PDF_Core implements PHPExcel_Writer_IWriter {
	/**
	 * Create a new PHPExcel_Writer_PDF
	 *
	 * @param 	PHPExcel	$phpExcel	PHPExcel object
	 */
    const INCH_FACTOR = 25.4;
	public function __construct(PHPExcel $phpExcel) {
		parent::__construct($phpExcel);
	}

	/**
	 * Save PHPExcel to file
	 *
	 * @param 	string 		$pFileName
	 * @throws 	Exception
	 */
	public function save($pFilename = null) {
		// garbage collect
		$this->_phpExcel->garbageCollect();

		$saveArrayReturnType = PHPExcel_Calculation::getArrayReturnType();
		PHPExcel_Calculation::setArrayReturnType(PHPExcel_Calculation::RETURN_ARRAY_AS_VALUE);

		// Open file
		$fileHandle = fopen($pFilename, 'w');
		if ($fileHandle === false) {
			throw new Exception("Could not open file $pFilename for writing.");
		}

		// Set PDF
		$this->_isPdf = true;
		// Build CSS
		$this->buildCSS(true);

		// Default PDF paper size
		$paperSize = 'LETTER';	//	Letter	(8.5 in. by 11 in.)

		// Check for paper size and page orientation
		if (is_null($this->getSheetIndex())) {
            $sheetIndex = 0;

		} else {
            $sheetIndex = $this->getSheetIndex();

		}
        $sheet = $this->_phpExcel->getSheet($sheetIndex);
        $orientation = ($sheet->getPageSetup()->getOrientation() == PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE) ? 'L' : 'P';
        $printPaperSize = $sheet->getPageSetup()->getPaperSize();
        $printMargins = $sheet->getPageMargins();
        $font = $sheet->getDefaultStyle()->getFont();

		$this->setOrientation($orientation);

		//	Override Page Orientation
		if (!is_null($this->getOrientation())) {
			$orientation = ($this->getOrientation() == PHPExcel_Worksheet_PageSetup::ORIENTATION_DEFAULT) ?
				PHPExcel_Worksheet_PageSetup::ORIENTATION_PORTRAIT : $this->getOrientation();
		}
		$orientation = strtoupper($orientation);

		//	Override Paper Size
		if (!is_null($this->getPaperSize())) {
			$printPaperSize = $this->getPaperSize();
		}

		if (isset(self::$_paperSizes[$printPaperSize])) {
			$paperSize = self::$_paperSizes[$printPaperSize];
		}

		// Create PDF
        //excel storage data in inches
        $inchFactor = self::INCH_FACTOR;

		$pdf = new mpdf(
            '', //mode
            'A4', // page standart
           $font->getSize(), // font size - inherit from excel
           $font->getName(),// default font - inherit from excel,

           $printMargins->getLeft()*$inchFactor, // milimetres
           $printMargins->getRight()*$inchFactor,
           $printMargins->getTop()*$inchFactor,
           $printMargins->getBottom()*$inchFactor,
           $printMargins->getHeader()*$inchFactor,
           $printMargins->getFooter()*$inchFactor
        );
		$pdf->_setPageSize(strtoupper($paperSize), $orientation);
        $pdf->DefOrientation = $orientation;
		// Document info
		$pdf->SetTitle($this->_phpExcel->getProperties()->getTitle());
		$pdf->SetAuthor($this->_phpExcel->getProperties()->getCreator());
		$pdf->SetSubject($this->_phpExcel->getProperties()->getSubject());
		$pdf->SetKeywords($this->_phpExcel->getProperties()->getKeywords());
		$pdf->SetCreator($this->_phpExcel->getProperties()->getCreator());

        $html = $this->generateHTMLHeader(TRUE) .
            $this->generateSheetData() .
            $this->generateHTMLFooter();

        $pdf->WriteHTML(
            $html
		);
        
		// Write to file
		fwrite($fileHandle, $pdf->Output('','S'));

		// Close file
		fclose($fileHandle);

		PHPExcel_Calculation::setArrayReturnType($saveArrayReturnType);
	}

}
