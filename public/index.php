<?php
require_once __DIR__ . '/../src/XlsxReaderBasic.php';
use ReaderXlsx\XlsxReaderBasic;

$file = __DIR__ . '/nps.xlsx';
if (!file_exists($file)) {
    echo "Arquivo nps.xlsx não encontrado.";
    exit;
}

$reader = new XlsxReaderBasic($file);
$headers = $reader->getHeaders();

echo '<h2>Headers do arquivo nps.xlsx:</h2>';
echo '<pre>' . htmlspecialchars(print_r($headers, true)) . '</pre>';

foreach ($reader->getRows() as $row) {
    echo '<pre>' . htmlspecialchars(print_r($row, true)) . '</pre>';
}