<?php

namespace Itszun\Importer;

use Carbon\Carbon;
use Closure;
use PhpOffice\PhpSpreadsheet\IOFactory;

use function Laravel\Prompts\info;
use function Laravel\Prompts\progress;

/**
 * Minimal service for handling import process from excel to array
 */
class Importer
{
    protected $source = "";
    protected Closure $transformer;
    protected bool $deleteAfterFinished = false;

    /**
     * Set file source
     *
     * @param string $source
     * @return static
     */
    public function source(string $source)
    {
        $this->source = $source;
        return $this;
    }

    /**
     * Helper for converting date
     *
     * @param string $data
     * @return string|null
     */
    public static function convertDate($data)
    {
        // Detect 21/06/2003
        if (preg_match('/^[0-9]{1,2}\/[0-9]{1,2}\/[0-9]{4}$/', $data)) {
            return Carbon::createFromFormat("d/m/Y", $data)->format('Y-m-d');
        }
        // Detect 21-06-2003
        if (preg_match('/^[0-9]{1,2}-[0-9]{1,2}-[0-9]{4}$/', $data)) {
            return Carbon::createFromFormat("d-m-Y", $data)->format('Y-m-d');
        }
        // Detect 21/06/03
        if (preg_match('/^[0-9]{1,2}\/[0-9]{1,2}\/[0-9]{2}$/', $data)) {
            return Carbon::createFromFormat("d/m/Y", $data)->format('Y-m-d');
        }
        // Detect 21-06-03
        if (preg_match('/^[0-9]{1,2}-[0-9]{1,2}-[0-9]{2}$/', $data)) {
            return Carbon::createFromFormat("d-m-Y", $data)->format('Y-m-d');
        }
        return null;
    }

    /**
     * Set function callback/handler after excel to array 
     *
     * @param Closure|callable $transformer
     * @return static
     */
    public function transformer(Closure|callable $transformer)
    {
        $this->transformer = is_callable($transformer) ? Closure::fromCallable($transformer) : $transformer;
        return $this;
    }

    /**
     * Set deletion after finished process
     *
     * @param boolean $delete
     * @return static
     */
    public function deleteAfterFinished($delete = true)
    {
        $this->deleteAfterFinished = $delete;
        return $this;
    }

    /**
     * Start converting excel to array and use transformer, deleting after finished accordingly
     *
     * @return dynamic
     */
    public function process()
    {
        info("PROCESSING");
        $data = static::excelToArray($this->source);
        $result = $this->transformer->call($this, $data);
        if ($this->deleteAfterFinished && file_exists($this->source)) {
            unlink($this->source);
        }
        return $result;
    }

    /**
     * Converting Excel to Array
     * 
     * Extending from PhpOffice\PhpSpreadsheet\IOFactory library
     * 
     * @param string $path File Path
     * 
     * @return array
     */
    public static function excelToArray(string $path)
    {
        ini_set('memory_limit', '-1');
        set_time_limit(0);
        info("Identifying Source");
        $check_format = IOFactory::identify($path);
        $reader = IOFactory::createReader($check_format);
        info("Loading...");
        $str_time = microtime(true);
        $object = $reader->load($path);
        info("Finished Loading => " . (microtime(true) - $str_time) / 60 . "m");

        $sh = 0;
        $data = [];
        $iterators = $object->getWorksheetIterator();
        $str_time = microtime(true);
        info("Start Converting Data ...");
        foreach ($iterators as $worksheet) {
            if ($sh < 1) {
                $highestRow = $worksheet->getHighestRow();
                $highestColumn = $worksheet->getHighestColumn();

                for ($row = 1; $row <= $highestRow; $row++) {
                    $child_data = [];
                    $isEmptyRow = true; // cek if empety

                    for ($col = 0; $col <= static::convertChrToNum($highestColumn); $col++) {
                        $value = trim($worksheet->getCell([$col + 1, $row])->getFormattedValue());
                        // $child_data[$col+1] = trim($worksheet->getCell([$col + 1, $row])->getFormattedValue());
                        $child_data[$col + 1] = $value;

                        // If row tidak kosong
                        if (!empty($value)) {
                            $isEmptyRow = false;
                        }
                    }

                    // uwu
                    if (!$isEmptyRow) {
                        $data[$row] = $child_data;
                    }
                }
            }
            $sh++;
        }
        info("Converting Data Finished => " . (microtime(true) - $str_time) / 60 . "m");
        return $data;
    }

    /**
     * Converting alpha into number
     * 
     * For example: A = 1, AA = 27
     *
     * @param string $a
     * @return int
     */
    public static function convertChrToNum($a)
    {
        $r = 0;
        $l = strlen($a);
        for ($i = 0; $i < $l; $i++) {
            $r += pow(26, $i) * (ord($a[$l - $i - 1]) - 0x40);
        }
        return $r;
    }

    /**
     * Converting number into alpha
     * 
     * For example: 1 = 1, 27 = AA
     *
     * @param int $n
     * @return string
     */
    public static function convertNumToChr($n)
    {
        $r = '';
        for ($i = 1; $n >= 0 && $i < 10; $i++) {
            $r = chr(0x41 + ($n % pow(26, $i) / pow(26, $i - 1))) . $r;
            $n -= pow(26, $i);
        }
        return $r;
    }
}
