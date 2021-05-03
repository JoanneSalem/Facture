<?php

namespace App\Util;

use PhpOffice\PhpSpreadsheet\Reader\Xls\Color;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Color as StyleColor;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx as WriterXlsx;
use PhpOffice\PhpSpreadsheet\Style\Border;

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
                    'Produit1'=> array(
                        'libelle'=>'lait', 
                        'quantité'=>50, 
                        'prix'=>700,
                        'total'=>2000
                    ),
                    'Produit2'=> array(
                        'libelle'=>'sucre', 
                        'quantité'=>50, 
                        'prix'=>200,
                        'total'=>3000
                    ),
                    'Produit3'=> array(
                        'libelle'=>'poudre', 
                        'quantité'=>20, 
                        'prix'=>1000,
                        'total'=>7000
                    ),    
                    'Produit4'=> array(
                        'libelle'=>'soude', 
                        'quantité'=>20, 
                        'prix'=>1000,
                        'total'=>8000
                    ), 
                    'Produit5'=> array(
                        'libelle'=>'chausette', 
                        'quantité'=>20, 
                        'prix'=>1000,
                        'total'=>6000
                    ),    
                    'Produit6'=> array(
                        'libelle'=>'pipette', 
                        'quantité'=>20, 
                        'prix'=>1000,
                        'total'=>7000
                    ),       
            ),
            'Montant total'=>3000
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
                        'prix'=>700,
                        'total'=>1000
                    ),
                    'Produit2'=> array(
                        'libelle'=>'glace', 
                        'quantité'=>50, 
                        'prix'=>200,
                        'total'=>1000
                    ),
                    'Produit3'=> array(
                        'libelle'=>'jus', 
                        'quantité'=>20, 
                        'prix'=>1000,
                        'total'=>1000
                    ),   
                 
            ),
            'Montant total'=>5000
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
                        'prix'=>700,
                        'total'=>1000
                    ),
                    'Produit2'=> array(
                        'libelle'=>'Pomme', 
                        'quantité'=>10, 
                        'prix'=>200,
                        'total'=>1000
                    ),
                    'Produit3'=> array(
                        'libelle'=>'Haricots', 
                        'quantité'=>2, 
                        'prix'=>1000,
                        'total'=>1000
                    ),
                 
            ),
            'Montant total'=>10000
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
                        'prix'=>7000,
                        'total'=>1000
                    ),
                    'Produit2'=> array(
                        'libelle'=>'Chaussure', 
                        'quantité'=>10, 
                        'prix'=>2000,
                        'total'=>1000
                    ),
                    'Produit3'=> array(
                        'libelle'=>'Robe', 
                        'quantité'=>2, 
                        'prix'=>8000,
                        'total'=>1000
                    ),
                     
            ),
            'Montant total'=>20000
        );

        $listeCommande = array($facture1, $facture2, $facture3, $facture4);
        

        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();



        $spreadsheet->getActiveSheet()->getColumnDimension('A')->setAutoSize(true);
        #$spreadsheet->getActiveSheet()->getColumnDimension('B')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('C')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('D')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('E')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('F')->setAutoSize(true);


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
        $spreadsheet->getActiveSheet()->getStyle('C'.$i)->getFont()->setBold(true);
        
        $sheet->setCellValue('D'.$i, $valeur['Numéro de Reçu']);
        $sheet->setCellValue('C'.$i++, 'N° de Reçu :');
        $spreadsheet->getActiveSheet()->getStyle('C'.$i)->getFont()->setBold(true);


        $sheet->setCellValue('B'.$i, $valeur['Nom']);
        $spreadsheet->getActiveSheet()->getStyle('C'.$i)->getFont()->setBold(true);
        $sheet->setCellValue('C'.$i, 'Vente N° :');
        $sheet->setCellValue('D'.$i, $valeur['Numéro de vente']);
        $spreadsheet->getActiveSheet()->getStyle('A'.$i)->getFont()->setBold(true);
        $sheet->setCellValue('A'.$i++, 'Nom');
        

        $spreadsheet->getActiveSheet()->getStyle('A'.$i.':F'.$i)->getFill()
        ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
        ->getStartColor()->setARGB('CECECE');

        $sheet->setCellValue('A'.$i, 'Description');
        $sheet->setCellValue('D'.$i, 'Quantité');
        $spreadsheet->getActiveSheet()->getStyle('D'.$i)->getFont()->setBold(true);
        $sheet->setCellValue('E'.$i, 'Prix Unitaire');
        $spreadsheet->getActiveSheet()->getStyle('E'.$i)->getFont()->setBold(true);
        $sheet->setCellValue('F'.$i, 'Total');
        $spreadsheet->getActiveSheet()->getStyle('F'.$i)->getFont()->setBold(true);
        $spreadsheet->getActiveSheet()->getStyle('A'.$i++)->getFont()->setBold(true);
       
        
        
        $j=0;
        foreach($valeur['Ligne de commande'] as $commande){

                $sheet->setCellValue('A'.$i, $commande['libelle']);
                $sheet->setCellValue('D'.$i, $commande['quantité']);
                $sheet->setCellValue('E'.$i, $commande['prix']);
                $sheet->setCellValue('F'.$i, $commande['total']);      
                $i++;
                $j++;
                if ($j>=11){
                    break;
                };
        };

        for($k=1; 11-$j>$k; $k++){
            $sheet->setCellValue('F'.$i++, '0');
        };

        $sheet->setCellValue('F'.$i, $valeur['Montant total']);
        $sheet->setCellValue('E'.$i++, 'TOTAL :');
     
        
        $i+=2;
        };
        $sheet->getStyle("A1:F".$i)->getFont()->setSize(5);

        $writer = new WriterXlsx($spreadsheet);
        $writer->save('hello world.xlsx');

    }
}
