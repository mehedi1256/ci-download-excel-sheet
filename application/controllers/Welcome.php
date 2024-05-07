<?php
defined('BASEPATH') OR exit('No direct script access allowed');
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
class Welcome extends CI_Controller {

	/**
	 * Index Page for this controller.
	 *
	 * Maps to the following URL
	 * 		http://example.com/index.php/welcome
	 *	- or -
	 * 		http://example.com/index.php/welcome/index
	 *	- or -
	 * Since this controller is set as the default controller in
	 * config/routes.php, it's displayed at http://example.com/
	 *
	 * So any other public methods not prefixed with an underscore will
	 * map to /index.php/welcome/<method_name>
	 * @see https://codeigniter.com/userguide3/general/urls.html
	 */

	
	public function index() {
		$this->load->helper('url');
		$this->load->view('welcome_message');
	}


	public function download_excel_sheet()
	{
		$spreadsheet = new Spreadsheet();
		$sheet = $spreadsheet->getActiveSheet();
		
		foreach(range('A','F') as $coulumID) {
			$spreadsheet->getActiveSheet()->getColumnDimension($coulumID)->setAutosize(true);

		}
		$sheet->setCellValue('A1','ID');
		$sheet->setCellValue('B1','Name');
		$sheet->setCellValue('C1','EMAIL');
		$sheet->setCellValue('D1','MOBILE');
		$sheet->setCellValue('E1','CITY');
		$sheet->setCellValue('F1','COUNTRY');

		$users = $this->db->query("SELECT * FROM users")->result_array();
		$x=2; //start from row 2
		foreach($users as $row)
		{
			$sheet->setCellValue('A'.$x, $row['id']);
			$sheet->setCellValue('B'.$x, $row['username']);
			$sheet->setCellValue('C'.$x, $row['email']);
			$sheet->setCellValue('D'.$x, $row['mobile']);
			$sheet->setCellValue('E'.$x, $row['city']);
			$sheet->setCellValue('F'.$x, $row['country']);
			$x++;
		}

		$writer = new Xlsx($spreadsheet);
		$fileName='users_details_export2022.xlsx';
		//$writer->save($fileName);  //this is for save in folder


		/* for force download */
		header('Content-Type: appliction/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
		header('Content-Disposition: attachment; filename="'.$fileName.'"');
		$writer->save('php://output');
		/* force download end */
	}
}
