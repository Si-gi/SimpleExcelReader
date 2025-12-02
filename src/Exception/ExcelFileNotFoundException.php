<?php
declare(strict_types=1);

namespace Sigi\SimpleExcelReader\Exception;

class ExcelFileNotFoundException extends SimpleExcelReaderException
{
    public function __construct(string $filePath, ?\Throwable $previous = null)
    {
        parent::__construct(
            sprintf("Excel file '%s' not found", $filePath),
            404,
            $previous,
            $filePath
        );
    }
}