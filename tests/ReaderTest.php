<?php

declare(strict_types=1);

namespace Sigi\SimpleExcelReader\Tests;

use PHPUnit\Framework\TestCase;
use SebastianBergmann\Timer\Timer;
use Sigi\SimpleExcelReader\Reader;

class ReaderTest extends TestCase
{
    private string $excelPath;
    private Reader $reader;
    private Timer $timer;
    protected function setUp(): void
    {
        $this->excelPath = realpath(__DIR__."/ressources/xlsx_with_300k_rows_and_shared_strings.xlsx");


        $this->reader = new Reader(new \ZipArchive(), new \XMLReader());

        $this->timer = new Timer;

    }

    public function testOpen(): void
    {
        $open = $this->reader->open($this->excelPath);
        self::assertInstanceOf(Reader::class, $open);
    }

    public function testLoadRead(): void
    {
        $this->timer->start();

        $this->reader->open($this->excelPath);
        $this->reader->loadSheets();
        $this->reader->loadRelationships();
        $sheet =  $this->reader->getSheets()[0];
        $tempSheet = $this->reader->extractSheet($sheet);
        $this->reader->loadSheet($tempSheet);
        $this->reader->loadSharedString();
        $row = $this->reader->read([$this->reader, 'findValue'], ['xlsx--300000-3']);
        $duration = $this->timer->stop();
        
        self::assertIsArray($row);
        self::assertContains('xlsx--300000-1', $row);
    }

    public function testReaderReturnManyRow(): void
    {
        $this->timer->start();

        $this->reader->open($this->excelPath);
        $this->reader->loadSheets();
        $this->reader->loadRelationships();
        $sheet =  $this->reader->getSheets()[0];
        $tempSheet = $this->reader->extractSheet($sheet);
        $this->reader->loadSheet($tempSheet);
        $this->reader->loadSharedString();

        $rows = $this->reader->findManyRow([$this->reader, 'findValue'], ['xlsx--300000-3', 'xlsx--299999-3']);
        var_dump($rows);
        $duration = $this->timer->stop();
        self::assertTrue($duration->asSeconds() < 25);
        self::assertIsArray($rows);
        self::assertContains('xlsx--299999-1', $rows[0]);

        self::assertContains('xlsx--300000-1', $rows[1]);

    }
}
