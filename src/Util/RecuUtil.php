<?php

namespace App\Util;

use PhpOffice\PhpSpreadsheet\Reader\Xls\Color;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Color as StyleColor;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx as WriterXlsx;

class RecuUtil 
{
    
    public static function index()
    {
        
        
        $facture1 = array(
 
            'Date' => '13-02-2021', 
            'Numéro de Reçu' => 1, 
            'Numéro de vente' => 1,
            'Nom' => 'LeDuc',
            'Ligne de commande' => array(
                    'Produit'=> array(
                        'libelle'=>'lait', 
                        'quantité'=>50, 
                        'prix'=>700
                    ),
                    'Produit'=> array(
                        'libelle'=>'sucre', 
                        'quantité'=>50, 
                        'prix'=>200
                    ),
                    'Produit'=> array(
                        'libelle'=>'poudre', 
                        'quantité'=>20, 
                        'prix'=>1000
                    ),    
            'Montant total'=>2000
            ),

        );

        $facture2 = array(
            'Date' => '14-02-2021', 
            'Numéro de Reçu' => 2, 
            'Numéro de vente' => 2,
            'Nom' => 'Jean',
            'Ligne de commande' => array(
                
                    'Produit1'=> array(
                        'libelle'=>'lait', 
                        'quantité'=>50, 
                        'prix'=>700
                    ),
                    'Produit2'=> array(
                        'libelle'=>'glace', 
                        'quantité'=>50, 
                        'prix'=>200
                    ),
                    'Produit3'=> array(
                        'libelle'=>'jus', 
                        'quantité'=>20, 
                        'prix'=>1000
                    ),   
                 
                'Montant total'=>5000
            ),

        );
        
        $facture3 = array(
            'Date' => '15-02-2021', 
            'Numéro de Reçu' => 3, 
            'Numéro de vente' => 3,
            'Nom' => 'Jean',
            'Ligne de commande' => array(
                    'Produit1'=> array(
                        'libelle'=>'Saucisse', 
                        'quantité'=>5, 
                        'prix'=>700
                    ),
                    'Produit2'=> array(
                        'libelle'=>'Pomme', 
                        'quantité'=>10, 
                        'prix'=>200
                    ),
                    'Produit3'=> array(
                        'libelle'=>'Haricots', 
                        'quantité'=>2, 
                        'prix'=>1000
                    ),
                 
                'Montant total'=>10000
            ),

        );

        $facture4 = array(
            'Date' => '17-02-2021', 
            'Numéro de Reçu' => 4, 
            'Numéro de vente' => 4,
            'Nom' => 'Paul',
            'Ligne de commande' => array(
                    'Produit1'=> array(
                        'libelle'=>'Livre', 
                        'quantité'=>5, 
                        'prix'=>7000
                    ),
                    'Produit2'=> array(
                        'libelle'=>'Chaussure', 
                        'quantité'=>10, 
                        'prix'=>2000
                    ),
                    'Produit3'=> array(
                        'libelle'=>'Robe', 
                        'quantité'=>2, 
                        'prix'=>8000
                    ),
                    
                
                 
                'Montant total'=>20000
            ),

        );

        $listeCommande = array($facture1, $facture2, $facture3, $facture4);
        #$listeProduit = array($produit1,);

        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();


        $i=1;
        foreach($listeCommande as $valeur){

        $drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\Drawing();
        $drawing->setName('Logo');
        $drawing->setDescription('Logo');
        $drawing->setCoordinates('A'.$i);
        $drawing->setPath('C:/wamp64/www/recu/public/logo.jpg'); 
        $drawing->setHeight(46);
        $drawing->setWorksheet($spreadsheet->getActiveSheet());
       
        $sheet->setCellValue('D'.$i, $valeur['Date']);
        $sheet->setCellValue('C'.$i++, 'Date :');
        

        $sheet->setCellValue('C'.$i, 'N° de Reçu :');
        $sheet->setCellValue('D'.$i++, $valeur['Numéro de Reçu']);


        $sheet->setCellValue('B'.$i, $valeur['Nom']);
        $sheet->setCellValue('C'.$i, 'Vente N° :');
        $sheet->setCellValue('D'.$i, $valeur['Numéro de vente']);
        $sheet->setCellValue('A'.$i++, 'Nom');

        
        
        $sheet->setCellValue('D'.$i, 'Quantité');
        $sheet->setCellValue('E'.$i, 'Prix Unitaire');
        $sheet->setCellValue('F'.$i, 'Total');

        foreach($valeur['Ligne de commande'] as $commande){
            foreach($commande['Produit1'] as $produit){
                $sheet->setCellValue('A'.$i, $produit['libelle']);
                #$sheet->setCellValue('A'.$i, $commande['quantité']);
                #$sheet->setCellValue('A'.$i, $commande['prix']);
            };
            $i++;
        };


        $sheet->setCellValue('A'.$i++, 'Description');

        $sheet->setCellValue('F'.$i++, '0');
        $sheet->setCellValue('F'.$i++, '0');
        $sheet->setCellValue('F'.$i++, '0');
        $sheet->setCellValue('F'.$i++, '0');
        $sheet->setCellValue('F'.$i++, '0');
        $sheet->setCellValue('F'.$i++, '0');
        $sheet->setCellValue('F'.$i++, '0');
        $sheet->setCellValue('F'.$i++, '0');
        $sheet->setCellValue('F'.$i++, '0');
        $sheet->setCellValue('F'.$i++, '0');
        $sheet->setCellValue('F'.$i++, '0');
        $sheet->setCellValue('E'.$i++, 'TOTAL :');
        

        $i+=4;
        };
        


        $spreadsheet->getActiveSheet()->getStyle('A4:F4')->getFill()
        ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
        ->getStartColor()->setARGB('CECECE');



        $spreadsheet->getActiveSheet()->getColumnDimension('A')->setAutoSize(true);
        #$spreadsheet->getActiveSheet()->getColumnDimension('B')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('C')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('D')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('E')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('F')->setAutoSize(true);


        $spreadsheet->getActiveSheet()->getStyle('C1')->getFont()->setBold(true);
        $spreadsheet->getActiveSheet()->getStyle('A3')->getFont()->setBold(true);
        $spreadsheet->getActiveSheet()->getStyle('A4')->getFont()->setBold(true);
        $spreadsheet->getActiveSheet()->getStyle('C2')->getFont()->setBold(true);
        $spreadsheet->getActiveSheet()->getStyle('C3')->getFont()->setBold(true);
        $spreadsheet->getActiveSheet()->getStyle('D4')->getFont()->setBold(true);
        $spreadsheet->getActiveSheet()->getStyle('E4')->getFont()->setBold(true);
        $spreadsheet->getActiveSheet()->getStyle('E16')->getFont()->setBold(true);
        $spreadsheet->getActiveSheet()->getStyle('F4')->getFont()->setBold(true);


        $spreadsheet->getActiveSheet()->getStyle('C1')
        ->getBorders()->getTop()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
        $spreadsheet->getActiveSheet()->getStyle('C1')
        ->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
        $spreadsheet->getActiveSheet()->getStyle('C1')
        ->getBorders()->getLeft()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
        $spreadsheet->getActiveSheet()->getStyle('C1')
        ->getBorders()->getRight()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);

        $spreadsheet->getActiveSheet()->getStyle('C2')
            ->getBorders()->getTop()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
        $spreadsheet->getActiveSheet()->getStyle('C2')
            ->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
        $spreadsheet->getActiveSheet()->getStyle('C2')
            ->getBorders()->getLeft()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
        $spreadsheet->getActiveSheet()->getStyle('C2')
            ->getBorders()->getRight()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);

        $spreadsheet->getActiveSheet()->getStyle('C3')
            ->getBorders()->getTop()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
        $spreadsheet->getActiveSheet()->getStyle('C3')
            ->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
        $spreadsheet->getActiveSheet()->getStyle('C3')
            ->getBorders()->getLeft()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
        $spreadsheet->getActiveSheet()->getStyle('C3')
            ->getBorders()->getRight()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);    

        $spreadsheet->getActiveSheet()->getStyle('A4:F15')
            ->getBorders()->getTop()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
        $spreadsheet->getActiveSheet()->getStyle('A4:F15')
            ->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
        $spreadsheet->getActiveSheet()->getStyle('A4:F15')
            ->getBorders()->getLeft()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
        $spreadsheet->getActiveSheet()->getStyle('A4:F15')
            ->getBorders()->getRight()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);

    
        
        $spreadsheet->getActiveSheet()->mergeCells('A1:A2');
        $spreadsheet->getActiveSheet()->mergeCells('D1:F1');
        $spreadsheet->getActiveSheet()->mergeCells('D2:F2');
        $spreadsheet->getActiveSheet()->mergeCells('D3:F3');


        $writer = new WriterXlsx($spreadsheet);
        $writer->save('hello world.xlsx');

    }
}
