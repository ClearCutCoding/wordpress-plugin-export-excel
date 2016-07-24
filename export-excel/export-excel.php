<?php
/**
 * @package Export_Excel
 * @version 1.0
 */
/*
Plugin Name: Export Excel
Plugin URI: http://www.clearcutcoding.com/
Description: Allows frontend export of data tables to an Excel spreadsheet
Author: Etch Akhtar
Version: 1.0
Author URI: http://www.clearcutcoding.com/
*/

defined( 'ABSPATH' ) or die( 'Unknown resource' );

define( 'EXPORTEXCEL__PLUGIN_URL', plugin_dir_url( __FILE__ ) );
define( 'EXPORTEXCEL__PLUGIN_PATH', plugin_dir_path( __FILE__ ) );

add_shortcode( 'exportexcel', 'export_excel_button' );
wp_register_style( 'exportexcel', EXPORTEXCEL__PLUGIN_URL . 'styles.css' );
wp_enqueue_style('exportexcel');
add_action( 'wp', 'export_excel');
add_filter( 'query_vars', 'add_query_vars_filter' );


/**
 * Add custom query params to wordpress list
 *
 * @param string $vars wordpress query vars array
 * @return string the updated vars
 */
function add_query_vars_filter($vars) {
	$vars[] = "exportexcel";
	return $vars;
}


/**
 * Creates button html for exporting to excel
 *
 * @return string the button html
 */
function export_excel_button() {
	// Permalinks most not be plain as we're adding a parameter with ?
	return "
		<a href=\"" . get_permalink() . "?exportexcel=yes\" class=\"export-excel\">Export to Excel</a>
	";
}


/**
 * Creates excel spreadsheet for html tables requiring export
 *
 * @return binary the excel file
 */
function export_excel() {

	// Check if we requested an export
	$doExport = get_query_var( 'exportexcel', 'no' );
	if (!$doExport || $doExport !== 'yes') return;

	// Find all tables we want exported from the page
	$post = get_post();
	$content = $post->post_content;
	$tables = convertTablesArray($content);

	// Export to excel
	require_once EXPORTEXCEL__PLUGIN_PATH . 'PHPExcel/PHPExcel.php';
	require_once EXPORTEXCEL__PLUGIN_PATH . 'PHPExcel/PHPExcel/IOFactory.php';

	// Create new PHPExcel object
	$objPHPExcel = new \PHPExcel();
	$curSheet = -1;

	foreach ($tables as $sheetData)
	{
		$curSheet++;
		// Create sheet
		if ($curSheet != 0) $objPHPExcel->createSheet();

		$objPHPExcel->setActiveSheetIndex($curSheet);
		$objPHPExcel->getActiveSheet()->setTitle('Table ' . ($curSheet+1) );

		// Format data for excel
		$row = -1;
		$colCount = 0;
		$sheetArray = array();

		// Data
		foreach ($sheetData as $value)
		{
			$row++;
			$sheetArray[$row] = $value;

			// Get number of columns (maximum required if each row has different amount of td tags)
			if ($colCount ==0 || $colCount < count($value)) $colCount = count($value);
		}

		// Create sheet
		$objPHPExcel->getActiveSheet()->fromArray($sheetArray, NULL, 'A1');

		// Make header bold
		$first_letter = \PHPExcel_Cell::stringFromColumnIndex(0);
		$last_letter = \PHPExcel_Cell::stringFromColumnIndex($colCount-1);
		$header_range = "{$first_letter}1:{$last_letter}1";
		$objPHPExcel->getActiveSheet()->getStyle($header_range)->getFont()->setBold(true);

		// Resize columns
		foreach(range('A',\PHPExcel_Cell::stringFromColumnIndex($colCount-1)) as $columnID)
		{
		    $objPHPExcel->getActiveSheet()->getColumnDimension($columnID)
		        ->setAutoSize(true);
		}
	}

	// Redirect output to a client’s web browser (Excel5)
	header('Content-Type: application/vnd.ms-excel');
	header('Content-Disposition: attachment;filename=Export.xls');
	header('Cache-Control: max-age=0');
	$objWriter = \PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
	$objWriter->save('php://output');
	exit;
}


/**
 * Converts html data to a php array
 *
 * @param string $contents the post data
 * @return array html table data converted to a multi-dimensional array
 */
function convertTablesArray($contents)
{
    $DOM = new DOMDocument;
    $DOM->loadHTML($contents);

    $data = array();

	$tables = $DOM->getElementsByTagName('table');
	foreach($tables as $table) {
		// Make sure it has a flag to allow exporting
		if (!$table->hasAttribute('export') || $table->getAttribute('export') !== 'yes') continue;

		$tempTable = array();

	    foreach ($table->childNodes as $tr) {
	    	$tempTD = array();

	        foreach ($tr->childNodes as $td) {
	        	$tempTD[] = $td->nodeValue;
	        }

	        $tempTable[] = $tempTD;
	    }

	    $data[] = $tempTable;
	}

    return $data;
}


