<?php
declare(strict_types=1);

namespace Sigi\SimpleExcelReader\Exception;

use Exception;

class SimpleExcelReaderException extends Exception
{
    protected ?string $filePath = null;
    
    public function __construct(
        string $message = "",
        int $code = 0,
        ?\Throwable $previous = null,
        ?string $filePath = null
    ) {
        parent::__construct($message, $code, $previous);
        $this->filePath = $filePath;
    }
    
    public function getFilePath(): ?string
    {
        return $this->filePath;
    }
}