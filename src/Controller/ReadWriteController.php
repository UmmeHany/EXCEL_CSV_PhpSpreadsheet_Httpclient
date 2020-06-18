<?php

namespace App\Controller;

use Symfony\Bundle\FrameworkBundle\Controller\AbstractController;
use Symfony\Component\Routing\Annotation\Route;
use Symfony\Component\HttpFoundation\Response;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Reader\Csv as ReaderCsv;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx as ReaderXlsx;
use Symfony\Component\HttpClient\HttpClient;

class ReadWriteController extends AbstractController
{
    /**
     * @Route("/write", name="write")
     */
    public function index()
    {
        $spreadsheet = new Spreadsheet();
        
        /* @var $sheet \PhpOffice\PhpSpreadsheet\Writer\Xlsx\Worksheet */
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->setCellValue('A1', 'Hello World !');
        $sheet->setTitle("My First Worksheet");
        
        // Create your Office 2007 Excel (XLSX Format)
        $writer = new Xlsx($spreadsheet);
        
        // In this case, we want to write the file in the public directory
       // $publicDirectory = $this->get('kernel')->getProjectDir() . '/public';
        $publicDirectory =$this->getParameter('kernel.project_dir') . '/public';
        // e.g /var/www/project/public/my_first_excel_symfony4.xlsx
        $excelFilepath =  $publicDirectory . '/my_first_excel_symfony4.xlsx';
        
        // Create the file
        $writer->save($excelFilepath);
        
        // Return a text response to the browser saying that the excel was succesfully created
        return new Response("Excel generated succesfully");
    }

    
    public function readFile($filename)
    {
        //$filename='my_first_excel_symfony4.xlsx';
        $extension = pathinfo($filename, PATHINFO_EXTENSION);
    switch ($extension) {
        case 'xlsx':
            $reader = new ReaderXlsx();
            break;
        case 'csv':
            $reader = new ReaderCsv();
            break;
        default:
            throw new \Exception('Invalid extension');
    }
    $reader->setReadDataOnly(true);
    return $reader->load($filename);
    
        
    }

protected function createDataFromSpreadsheet($spreadsheet)
{
    $data_col = [];
    $data_name = [];
    $data_city = [];
    $index=0;
    foreach ($spreadsheet->getWorksheetIterator() as $worksheet) {
        foreach ($worksheet->getRowIterator() as $row) {
            $rowIndex = $row->getRowIndex();
            if ($rowIndex > 1) {

                $cellIterator = $row->getCellIterator();
                $cellIterator->setIterateOnlyExistingCells(false); // Loop over all cells, even if it is not set
            
                
                foreach ($cellIterator as $cell) {
                
                if ($rowIndex >1) {
                    //echo $rowIndex;
                    ++$index;
                    $data_name[$index] = $cell->getCalculatedValue();
                    //$data_city['city'][$rowIndex] = $cell->getCalculatedValue();
                    
                }
            }
            }

        }



        }
        return $data_name;
    }

    





/**
 * @Route("/import", name="import")
 */
public function importAction()
{
    $filename = $this->getParameter('kernel.project_dir') . '/public/parts.csv';
    if (!file_exists($filename)) {
        throw new \Exception('File does not exist');
    }

    $spreadsheet = $this->readFile($filename);
    $data = $this->createDataFromSpreadsheet($spreadsheet);

    /*return $this->render('index.html.twig', [
        'data' => $data,
    ]);*/
    //$data = 'bula';
    //var_dump($data["My First Worksheet"]["columnNames"][1]);
    //var_dump($data);
    var_dump($data);
    return new Response("imported succesfully");
    //var_dump($data_city);
   // return new Response('Excel read done' );
    /*$data=['a','b','c'];
    return $this->render('index.html.twig',[
            'data' => $data
            
        ]);*/
}


/**
 * @Route("/link", name="link")
 */

 public function linktest(){

    $client = HttpClient::create();
$response = $client->request('GET', 'https://api.github.com/repos/symfony/symfony-docs');

$statusCode = $response->getStatusCode();
// $statusCode = 200
$contentType = $response->getHeaders()['content-type'][0];
// $contentType = 'application/json'
$content = $response->getContent();
// $content = '{"id":521583, "name":"symfony-docs", ...}'
$content = $response->toArray();
// $content = ['id' => 521583, 'name' => 'symfony-docs', ...]

return new Response("link".$content['license']['name']);

 }
}
