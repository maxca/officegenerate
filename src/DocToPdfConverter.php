<?php

namespace Samark\OfficeGenerate;

use Exception;

/**
 * Trait DocToPdfConverter
 * @package Samark\OfficeGenerate
 * @author samark@chai
 */
trait DocToPdfConverter
{

    /**
     * @param string $filename
     * @param string $outputDir
     * @return string|null
     * @throws Exception
     */
    public function convertPdf($filename = '', $outputDir = '')
    {
        if (!is_dir($outputDir)) {
            mkdir($outputDir, 0777, true);
        }
        try {
            return shell_exec('export HOME=/tmp && libreoffice --headless -convert-to pdf --outdir ' . $outputDir . ' ' . $filename);
        } catch (Exception $exception) {
            throw new Exception('Export pdf fail');
        }
    }
}