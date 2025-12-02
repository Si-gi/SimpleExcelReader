<?php

declare(strict_types=1);

namespace Sigi\SimpleExcelReader\Exception;

class ExcelInvalidSheetException extends SimpleExcelReaderException {
    
    private ?string $reason = null;
    
    public function __construct(
        string $filePath, 
        ?string $reason = null,
        ?\Throwable $previous = null
    ) {
        $this->reason = $reason;
        
        $message = sprintf("The sheet '%s' is invalid", $filePath);
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