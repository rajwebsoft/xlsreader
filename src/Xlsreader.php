<?php
namespace Rajwebsoft\Xlsreader;
use Rajwebsoft\Xlsreader\Library\ITFXlsReader;
use Rajwebsoft\Xlsreader\Library\XLSXReader;

class Xlsreader {

    public function loadFile($filename)
    {
        $itfxlsreader = new ITFXlsReader($filename);
        return $itfxlsreader;
    }

    public function readFile($filename)
    {
        
        $xlsreader = new XLSXReader($filename);
        return $xlsreader;
    }
}