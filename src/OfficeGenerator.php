<?php

namespace Samark\OfficeGenerate;

use PhpOffice\PhpWord\TemplateProcessor;
/**
 * Class OfficeGenerator
 * @package Samark\OfficeGenerate
 * @author samark@chai
 */
class OfficeGenerator implements OfficeGeneratorInterface
{

    use DocToPdfConverter;

    /**
     * @var TemplateProcessor
     * PHPWord instance
     */
    protected $templateProcess;

    /** @var bool
     * set default need convert doc to pdf
     */
    protected $needPdf = true;

    /**
     * @var string
     * set template file destination
     */
    protected $template = '/public/001.docx';

    /** @var array
     * list of row with data
     */
    protected $rows = [
        'id' => [
            [
                'id'       => '1',
                'name'     => 'samark',
                'lastname' => 'chais',
                'prefix'   => 'Mr',
            ],
            [
                'id'       => '2',
                'name'     => 'Document',
                'lastname' => 'chais',
                'prefix'   => 'Mr',
            ],
            [
                'id'       => '3',
                'name'     => 'Google',
                'lastname' => 'chais',
                'prefix'   => 'Mr',
            ],
        ],
    ];

    /** @var array
     * list of replace data
     * by ${variable} value
     */
    protected $replaces = [
        'serverName'      => 'samark',
        'myReplacedValue' => 'google',
        'weekday'         => '138',
        'for'             => ' Supper informaiotn '
    ];

    /**
     * @var data
     * list of data for combine to document
     * rows and replace key
     */
    protected $data = [];

    /**
     * @var string $filename
     * set file name output
     */
    protected $filename = 'simple';

    /**
     * @var path file need to save
     * set path of output file
     */
    protected $pathSaveAs = '/public/maxca';

    /** @var string
     * set default export document version
     */
    protected static $docVersion = '.doc';

    /**
     * WordService constructor.
     * @throws \PhpOffice\PhpWord\Exception\CopyFileException
     * @throws \PhpOffice\PhpWord\Exception\CreateTemporaryFileException
     */
    public function __construct()
    {
        $this->template        = base_path($this->template);
        $this->pathSaveAs      = base_path($this->pathSaveAs);
        $this->templateProcess = new TemplateProcessor($this->template);
    }

    /**
     * processing data
     * 1.copy template to path save as
     * 2.make folder by user id
     * 3.gen file name with file type of file
     */
    public function process()
    {
        # check directory existing
        $this->checkDir();

        # copy file to destination path
        $filename = $this->pathSaveAs . '/' . $this->filename . self::$docVersion;
        copy($this->template, $filename);

        # merge data
        $this->mergeData();

        # replace value
        $this->replaceValue();

        # mapping value
        $this->mappingValue();

        # save override file word
        $this->save($filename);

        # convert pdf
        if ($this->needPdf === true) {
            $this->convertPdf($filename, $this->pathSaveAs . '/pdf/');
        }

        return [
            'doc' => $filename,
            'pdf' => $this->pathSaveAs . '/pdf/' . $this->filename . '.pdf',
        ];
    }

    /**
     * merge data row and replace
     */
    protected function mergeData()
    {
        $this->data['row']     = $this->rows;
        $this->data['replace'] = $this->replaces;
    }

    /**
     * @param array $data
     * @return $this
     */
    public function setData($data = array())
    {
        $this->data = $data;
        return $this;
    }

    /**
     * @return data
     */
    protected function getData()
    {
        return $this->data;
    }

    /**
     * @param array $rows
     * @return $this
     */
    public function setRows(array $rows = array())
    {
        $this->rows = $rows;
        return $this;
    }

    /**
     * @param array $replace
     * @return $this
     */
    public function setReplaces(array $replace = array())
    {
        $this->replaces = $replace;
        return $this;
    }

    /**
     * @param string $filename
     * set filename
     * @return $this
     */
    public function setFilename($filename = '')
    {
        $this->filename = $filename;
        return $this;
    }

    /**
     * checking directory existing
     * if not has directory continue to create directory
     */
    private function checkDir()
    {
        if (!is_dir($this->pathSaveAs)) {
            mkdir($this->pathSaveAs, 0777, true);
        }
    }

    /**
     * @param string $filename
     * @return void
     */
    public function save(string $filename)
    {
        $this->templateProcess->saveAs($filename);
    }

    /**
     * replacing value by mapping key
     * @return void
     */
    public function replaceValue()
    {
        foreach ($this->data['replace'] as $key => $value) {
            $this->templateProcess->setValue($key, $value);
        }
    }

    /**
     * @return void
     * @throws \PhpOffice\PhpWord\Exception\Exception
     */
    private function mappingValue()
    {
        foreach ($this->data['row'] as $filed => $row) {
            $this->templateProcess->cloneRow($filed, count($row));
            $this->cloneRow($row);
        }
    }

    /**
     * @param $row
     */
    private function cloneRow($row)
    {
        $i = 1;
        foreach ($row as $key => $columns) {
            foreach ($columns as $filed => $item) {
                $this->templateProcess->setValue($filed . '#' . $i, $item);
            }
            $i++;
        }
    }
}