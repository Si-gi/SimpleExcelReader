<?php
declare(strict_types=1);

namespace Sigi\SimpleExcelReader\Exception;

class ExcelInvalidFormatException extends SimpleExcelReaderException
{
    private ?string $expectedFormat = null;
    private ?string $detectedFormat = null;
    
    public function __construct(
        string $filePath,
        ?string $expectedFormat = null,
        ?string $detectedFormat = null,
        ?\Throwable $previous = null
    ) {
        $this->expectedFormat = $expectedFormat;
        $this->detectedFormat = $detectedFormat;
        
        $message = sprintf("Format invalide pour '%s'", $filePath);
        if ($expectedFormat && $detectedFormat) {
            $message .= sprintf(" (attendu: %s, détecté: %s)", $expectedFormat, $detectedFormat);
        }
        
        parent::__construct($message, 415, $previous, $filePath);
    }
    
    public function getExpectedFormat(): ?string
    {
        return $this->expectedFormat;
    }
    
    public function getDetectedFormat(): ?string
    {
        return $this->detectedFormat;
    }
}