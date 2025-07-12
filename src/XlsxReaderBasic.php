<?php
namespace ReaderXlsx;

use Generator;

class XlsxReaderBasic
{
    private $filePath;
    private $headers = [];
    private $headerRow = 0;
    private $doc = null;

    public function __construct(string $filePath)
    {
        $this->filePath = $filePath;
    }

    public function getHeaders(): array
    {
        if (!empty($this->headers)) {
            return $this->headers;
        }
        $this->loadHeaders();
        return $this->headers;
    }

    public function getRows($bAfterHeader = true): Generator
    {
        $this->loadHeaders();
        $rows = $this->doc->getElementsByTagName('row');
        for ($i = ($bAfterHeader ? $this->headerRow : 0); $i < $rows->length; $i++) {
            yield $this->parseRow($rows->item($i));
        }
    }

    private $sharedStrings = [];

    private function loadHeaders(): void
    {
        if (!empty($this->headers) && $this->doc !== null) {
            return;
        }
        $this->loadSheetContext();
        $this->headers = $this->parseHeadersFromSheetXml();
    }

    private function loadSheetContext(): void
    {
        if ($this->doc !== null && !empty($this->sharedStrings)) {
            return;
        }
        $zip = new \ZipArchive();
        $sharedStrings = [];
        if ($zip->open($this->filePath) === TRUE) {
            // Carrega sharedStrings.xml se existir
            $sharedXml = $zip->getFromName('xl/sharedStrings.xml');
            if ($sharedXml !== false) {
                $sharedStrings = $this->parseSharedStrings($sharedXml);
            }
            $xml = $zip->getFromName('xl/worksheets/sheet1.xml');
            if ($xml !== false) {
                $this->doc = new \DOMDocument();
                $this->doc->loadXML($xml);
                $this->sharedStrings = $sharedStrings;
            }
            $zip->close();
        }
    }

    private function parseHeadersFromSheetXml(): array
    {
        $rows = $this->doc->getElementsByTagName('row');
        if ($rows->length > 0) {
            for ($i = 0; $i < $rows->length; $i++) {
                if (count(array_filter($this->parseRow($rows->item($i) ))) > 2) {
                    $this->headerRow = $i;
                    return $this->parseRow($rows->item($i));
                }
            }
            //return $this->parseRow($rows->item(0));
        }
        return [];
    }

    private function parseRow($rowNode): array
    {
        $data = [];
        foreach ($rowNode->getElementsByTagName('c') as $cell) {
            $value = '';
            $type = $cell->getAttribute('t');
            $style = $cell->getAttribute('s');
            foreach ($cell->getElementsByTagName('v') as $v) {
                $rawValue = $v->nodeValue;
                if ($type === 's' && isset($this->sharedStrings[(int) $rawValue])) {
                    $value = $this->sharedStrings[(int) $rawValue];
                } elseif ($type === 'd') {
                    // Tipo data ISO 8601
                    $value = date('Y-m-d', strtotime($rawValue));
                } elseif ($type === '' && $this->isExcelDate($rawValue, $style)) {
                    // Data serial Excel
                    $value = $this->excelDateToString($rawValue);
                } else {
                    $value = $rawValue;
                }
            }
            $data[] = $value;
        }
        return $data;
    }

    // Detecta se o valor é uma data serial do Excel (heurística simples)
    private function isExcelDate($value, $style): bool
    {
        // Pode ser melhorado: aqui só verifica se é numérico e plausível como data Excel
        return is_numeric($value) && $value > 20000 && $value < 80000;
    }

    // Converte data serial Excel para string (YYYY-MM-DD)
    private function excelDateToString($serial): string
    {
        $unix = ($serial - 25569) * 86400;
        return gmdate('Y-m-d', $unix);
    }

    private function parseSharedStrings($xml): array
    {
        $strings = [];
        $doc = new \DOMDocument();
        $doc->loadXML($xml);
        $siList = $doc->getElementsByTagName('si');
        foreach ($siList as $si) {
            $text = '';
            foreach ($si->getElementsByTagName('t') as $t) {
                $preserve = $t->getAttribute('xml:space') === 'preserve';
                $tVal = $t->nodeValue;
                if ($preserve) {
                    $text .= $tVal;
                } else {
                    $text .= trim($tVal);
                }
            }
            $strings[] = $text;
        }
        return $strings;
    }
}
