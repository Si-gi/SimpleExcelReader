<?php

declare(strict_types=1);
namespace SiGi\SimpleExcelReader\Exception;

use Sigi\SimpleExcelReader\Exception\SimpleExcelReaderException;

class ExcelNotReadableException extends SimpleExcelReaderException
{
    private ?string $reason = null;
    
    public function __construct(
        string $filePath, 
        ?string $reason = null,
        ?\Throwable $previous = null
    ) {
        $this->reason = $reason;
        
        $message = sprintf("The file '%s' is not readeable", $filePath);
        if ($reason) {
            $message .= sprintf(" : %s", $reason);
        }
        
        parent::__construct($message, 403, $previous, $filePath);
    }
    
    public function getReason(): ?string
    {
        return $this->reason;
    }
}