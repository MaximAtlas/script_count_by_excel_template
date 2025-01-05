<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$spreadsheet = IOFactory::load('template.xlsx');

$sheet = $spreadsheet->getActiveSheet();

$data = $sheet->toArray();
$trimmedData = array_slice($data, 2);

$peopleCount = $data[3][1];
$peoples = array_fill_keys(explode('|', $data[4][1]), []);

$startPosition = 9;

for ($i = $peopleCount; $i > 0; $i--) {
    $currentMan = $data[0][$startPosition];
    foreach ($trimmedData as $row) {
        $value = $row[$startPosition];
        $debtors = $row[$startPosition + 2];
        if ($debtors == 'все') {
            $individualDebt = $value / $peopleCount;
            foreach ($peoples as $debtorName => $debts) {
                if ($debtorName !== $currentMan) {
                    $peoples[$currentMan][$debtorName] = ($peoples[$currentMan][$debtorName] ?? 0) + $individualDebt;
                }
            }
        } elseif (!empty($debtors)) {
            $debtorsAr = explode('|', $debtors);
            $individualDebt = (int)$value / count($debtorsAr);
            foreach ($debtorsAr as $debtorName) {
                if ($debtorName !== $currentMan) {

                    $peoples[$currentMan][$debtorName] = ($peoples[$currentMan][$debtorName] ?? 0) + $individualDebt;
                }
            }
        }
    }
    $startPosition += 4;
}

$finalDebts = [];

foreach ($peoples as $creditor => $debts) {
    foreach ($debts as $debtor => $amount) {
        if (isset($peoples[$debtor][$creditor])) {
            $reciprocal = $peoples[$debtor][$creditor];
            if ($amount > $reciprocal) {
                $finalDebts[$creditor][$debtor] = $amount - $reciprocal;
                unset($peoples[$debtor][$creditor]);
            } elseif ($amount < $reciprocal) {
                $finalDebts[$debtor][$creditor] = $reciprocal - $amount;
                unset($peoples[$debtor][$creditor]);
            }
        } else {
            $finalDebts[$creditor][$debtor] = $amount;
        }
    }
}

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();

$sheet->setCellValue('A1', 'Кто должен');
$sheet->setCellValue('B1', 'Кому должен');
$sheet->setCellValue('C1', 'Сколько должен');

$row = 2;
foreach ($finalDebts as $creditor => $debts) {
    foreach ($debts as $debtor => $amount) {
        $sheet->setCellValue("A$row", $debtor);
        $sheet->setCellValue("B$row", $creditor);
        $sheet->setCellValue("C$row", $amount);
        $row++;
    }
}

$writer = new Xlsx($spreadsheet);
$writer->save('final_debts.xlsx');
dump('true');
