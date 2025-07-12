<?php
require_once __DIR__ . '/../vendor/autoload.php';
use ReaderXlsx\XlsxReaderBasic;
use ReaderXlsx\XlsxToJson;

$file = __DIR__ . '/CATMAT-parcial-2025.xlsx';
if (!file_exists($file)) {
    echo "Arquivo nps.xlsx não encontrado.";
    exit;
}

$reader = new XlsxReaderBasic($file);
header('Content-Type: application/json; charset=utf-8');
echo htmlspecialchars((new XlsxToJson($reader))->toJson(
    [
        "grupo" => ["Código do Grupo","Nome do Grupo"],
        "classe" => ["Código da Classe", "Nome da Classe"],
        "subclasse" => ["Código do PDM", "Nome do PDM"]
    ]), true);