<?php

namespace Samark\OfficeGenerate;

/**
 * Interface OfficeGeneratorInterface
 * @package Samark\OfficeGenerate
 * @author samark@chai
 */
interface OfficeGeneratorInterface
{

    /**
     * @param string $filename
     * @return mixed
     */
    public function save($filename);

    /**
     * @param array $rows
     * @return mixed
     */
    public function setRows(array $rows = array());

    /**
     * @param array $replaces
     * @return mixed
     */
    public function setReplaces(array $replaces = array());
}