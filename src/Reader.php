<?php

declare(strict_types=1);

namespace Sigi\SimpleExcelReader;

use Exception;
use XMLReader;
use ZipArchive;
use RuntimeException;
use Sigi\SimpleExcelReader\Sheet;
use Sigi\SimpleExcelReader\Exception\ExcelNotReadableException;
use Sigi\SimpleExcelReader\Exception\ExcelFileNotFoundException;


class Reader
{
    private array $sharedStrings = [];
    private string $filePath;
    private array $headers = [];
    private array $sheets = [];
    private array $relations = [];

    private string $prefix = "";
    public function __construct(private ZipArchive $zipArchive, private XMLReader $reader)
    {
    }

    public function setPrefix(string $prefix): static
    {
        $this->prefix = $prefix;
        return $this;
    }

    public function getSheets(): array
    {
        return $this->sheets;
    }
    public function open(string $filePath): static
    {
        if (!file_exists($filePath)) {
            throw new ExcelFileNotFoundException("Fichier SIM non trouvé");
        }


        $open = $this->zipArchive->open($filePath);
        if (!$open) {
            throw new ExcelNotReadableException("Fichier zip non lisible");
        }
        $this->filePath = realpath($filePath);
        return $this;
    }

    // The workbook contain sheat data ( names, ID, etc)
    public function loadSheets(): static
    {
        $workbookXml = $this->zipArchive->getFromName('xl/workbook.xml');
        if (!$workbookXml) {
            throw new Exception("WorkBook XML invalide");
        }

        $workbookReader = new XMLReader();
        $workbookReader->XML($workbookXml);

        while ($workbookReader->read()) {
            if ($this->readingNodeType($workbookReader, XMLReader::ELEMENT) && $this->readingNodeName($workbookReader, $this->prefix."sheet")) {
                $sheetName = $workbookReader->getAttribute('name');
                $rId = $workbookReader->getAttribute('r:id');
                $sheetIdAttr = $workbookReader->getAttribute('sheetId');

                $this->sheets[] = new Sheet($sheetName, $sheetIdAttr, $rId);

            }
        }
        $workbookReader->close();
        return $this;
    }

    public function readXML(XMLReader $reader, string $xml): bool
    {
        $reader->XML($xml);
        return $reader->read();
    }
    public function readingNodeType(XMLReader $reader, int $nodeType): bool
    {
        return $reader->nodeType === $nodeType;
    }

    public function readingNodeName(XMLReader $reader, string $nodeName): bool
    {
        return $reader->name === $nodeName;
    }

    public function getSheetByName(string $sheetName): ?Sheet
    {
        foreach ($this->sheets as $sheet) {
            if ($sheet->name === $sheetName) {
                return $sheet;
            }
        }
        return null;
    }

    public function extractSheet(Sheet $sheet): string
    {
        if (!isset($this->relations[$sheet->rId])) {
            throw new Exception("Relation introuvable pour la feuille {$sheet->name}");
        }
        
        $internalPath = $this->relations[$sheet->rId];
        $internalPath = ltrim($internalPath, '/');
        

        $zipFile = str_replace('\\', '/', $this->filePath);
        
        if (!str_starts_with($internalPath, 'xl/')) {
            $internalPath = 'xl/' . $internalPath;
        }
        
        $zipPath = "zip://{$zipFile}#{$internalPath}";
        
        $tempFile = sys_get_temp_dir() . '/sheet_temp_' . uniqid() . '.xml';
        
        if (!copy($zipPath, $tempFile)) {
            $error = error_get_last();
            throw new Exception(
                "Impossible d'extraire la feuille Excel via $zipPath. " .
                "Erreur: " . ($error['message'] ?? 'inconnue')
            );
        }
        
        return $tempFile;
    }
    public function loadSharedString(): static
    {
        $sharedStringsXml = $this->zipArchive->getFromName('xl/sharedStrings.xml');
        if (!$sharedStringsXml) {
            return $this;
        }
        $ssReader = new XMLReader();
        $ssReader->XML($sharedStringsXml);

        while ($ssReader->read()) {
            if ($ssReader->nodeType == XMLReader::ELEMENT && ($ssReader->name == $this->prefix.'t')) {
                $this->sharedStrings[] = $ssReader->readString();
            }
        }
        $ssReader->close();

        return $this;
    }

    public function loadSheet(string $filePath): static
    {

        $this->reader->open($filePath);
        return $this;
    }
    // should we add a trigger to stop the reading ?
    public function findManyRow(?callable $callback, ?array $research = [], ... $args): array
    {
        $rows = [];
        while ($this->reader->read()) {

            // we seek <row> elements
            if ($this->reader->nodeType == XMLReader::ELEMENT &&
                str_contains($this->reader->name, 'row')
            ) {
                $rowXml = $this->reader->readOuterXML();

                if ($callback !== null && ($row = $callback($rowXml, $research, $args)) !== null  && !empty($row)) {
                    $rows[] = $row;
                }
            }
        }
        return $rows;
    }

    // The main method. We pass callback in order to get what we want
    public function read(?callable $callback, ?array $research = [], ... $args): ?array
    {

        while ($this->reader->read()) {

            // we seek <row> elements
            if ($this->reader->nodeType == XMLReader::ELEMENT &&
                str_contains($this->reader->name, 'row')
            ) {
                $rowXml = $this->reader->readOuterXML();

                if ($callback !== null && ($row = $callback($rowXml, $research, $args)) !== null  && !empty($row)) {
                    return $row;
                }
            }
        }
        return null;
    }

    public function findHeader(string $subject): ?array
    {
        $headers = $this->findValue($subject, ["/.*?/"]);
        if ($headers === null) {
            throw new RuntimeException("Header non trouver");
        }
        if (empty($headers)) {
            throw new RuntimeException("Header vide");
        }
        $this->headers = $headers;
        return $headers;
    }
    public function findWithExcludedValues(string $subject, array $research, array $excluded): ?array
    {
        $cells = $this->findValue($subject, $research);
        if (empty($cells) || $cells === null) {
            return null;
        }
        if (in_array($excluded, $cells, true)) {
            return null;
        }
        return $cells;
    }
    // Should we split by type of research ? and maybe adding strategy ?
    public function findValue(string $subject, array $research): ?array
    {
        
        $values = $this->extractValuesFromRowXml($subject);
        if (empty($values)) {
            return null;
        }

        foreach ($values as $value) {

            foreach ($research as $rule) {
                $isRegex = $this->isRegex($rule);
                if (($isRegex && preg_match($rule, $value)) || (!$isRegex && $value === (string)$rule)) {
                    return $values;
                }
            }
        }
        return null;
    }

    private function isRegex(string $pattern): bool
    {
        return strlen($pattern) > 2 &&
            $pattern[0] === '/' &&
            strrpos($pattern, '/') !== 0;
    }

    public function mergeHeadersAndRow(array $cells): array
    {
        if (count($cells) < count($this->headers)) {
            $cells = array_pad($cells, count($this->headers), null);
        }

        // truncate if a row contains more cells
        if (count($cells) > count($this->headers)) {
            $cells = array_slice($cells, 0, count($this->headers));
        }

        return array_combine($this->headers, $cells);
    }

    public function loadRelationships(): void
    {
        $relsXml = $this->zipArchive->getFromName('xl/_rels/workbook.xml.rels');
        $rels = simplexml_load_string($relsXml);
        $rels->registerXPathNamespace('rel', 'http://schemas.openxmlformats.org/package/2006/relationships');

        foreach ($rels->Relationship as $rel) {
            $id = (string) $rel['Id'];
            $target = (string) $rel['Target'];

            $this->relations[$id] = $target;
        }
    }

    /**
    * Return list of cells or null
    */
    private function extractValuesFromRowXml(string $rowXml): ?array
    {
        //get of cell
        preg_match_all('/<c[^>]*r="([A-Z]+)(\d+)"[^>]*>(.*?)<\/c>/s', $rowXml, $cells, PREG_SET_ORDER);
        if (empty($cells)) {
            return null;
        }
        $rowValues = [];
        foreach ($cells as $cell) {

            $content   = $cell[3];

            if (preg_match('/<v>([^<]+)<\/v>/', $content, $vMatch)) {
                $value = $vMatch[1];
            } else {
                $value = "";
            }
        

            // Check if cell is a sharedString
            $isSharedString = str_contains($cell[0], 't="s"');
            if ($isSharedString) {
                $value = $this->cellValueFromSharedString($value);
            }
            $rowValues[] = $value;
        }

        return $rowValues;
    }

    private function cellValueFromSharedString(mixed $value): string
    {
        return $this->sharedStrings[(int)$value] ?? '';
    }

}
