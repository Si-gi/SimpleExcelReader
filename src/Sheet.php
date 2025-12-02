<?php

declare(strict_types=1);

namespace Sigi\SimpleExcelReader;

class Sheet{
    public function __construct(public string $name, public string $id, public string $rId, public ?string $path = null)
    {
    }

}