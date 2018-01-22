<?php

namespace chilipepper987;

/**
@author Seth Cohen
Simple PHP XLSX Writer v1.0
*/
class XLSXWriter {

    /**
     * Base64 encode of a "blank" xlsx file that we can inject the generated xml into. don't touch this
     */
    const TEMPLATE = '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';

    private $_tempDirectory;
    private $_filename, $_csv, $_tempFolderName, $_error, $_exec;
    private $_inputFile, $_inputString, $_inputArray;
    private $_freezePanes = false;

    public function __construct() {
        $this->_tempDirectoryBase = sys_get_temp_dir();
        $this->_font = $this->_border = $this->_fill = array();
    }

    /**
     * @param array $array The array to read in
     * @param bool $associative Assumes the array is an associative array, and makes the first row of the sheet (column headings) be array_keys($array[0]). Set to false if you don't want that, ie, if the first row in your input array is the column headings
     * @return bool
     */
    public function readArray($array, $associative = true) {
        if (is_array($array)) {
            $this->_inputArray = $array;
            if ($associative) {
                //make the first row, column headings, be the key of the first row of the input
                array_unshift($this->_inputArray, array_keys($array[0]));
            }
            return true;
        } else {
            return false;
        }
    }

    /**
     * Freeze the top n rows or first n columns of the spreedsheet.
     * @param string $type "top" or "left". Set to "off" to unfreeze panes.
     * @param int $n The number of rows (or columns) to freeze
     */
    public function freezePanes($type, $n, $n2 = 0) {
        if (in_array($type, ["top", "left", "top-left", "off"])) {
            if ($type === "off") {
                $this->_freezePanes = false;
            } else {
                $this->_freezePanes = ['type' => $type, 'n' => $n, 'n2' => $n2];
            }
        }
        return $this->_freezePanes;
    }

    private function _addPaneElement(XMLWriter $xml) {
        $splitAttributes = [];
        if (!empty($this->_freezePanes)) {
            if ($this->_freezePanes['type'] === "top") {
                //ySplit
                $splitAttributes[] = ["ySplit" => "" . $this->_freezePanes['n']];
                $topLeftCell = "A" . ($this->_freezePanes['n'] + 1);
                $activePane = "bottomLeft";
            } elseif ($this->_freezePanes["type"] === "left") {
                //xSplit
                $splitAttribute[] = ["xSplit" => "" . $this->_freezePanes['n']];
                $topLeftCell = $this->_num2alpha($this->_freezePanes['n'] + 1) . "1";
                $activePane = "topRight";
            } elseif ($this->_freezePanes["type"] === "top-left") {
                //top-left gets y AND x split
                $splitAttributes[] = ["ySplit" => "" . $this->_freezePanes['n']];
                $splitAttributes[] = ["xSplit" => "" . $this->_freezePanes['n2']];
                $topLeftCell = $this->_num2alpha($this->_freezePanes['n'] + 1) . ($this->_freezePanes['n2'] + 1);
                $activePane = "bottomRight";
            } else {
                return false;
            }
            //open pane
            $xml->startElement("pane");
            foreach ($splitAttributes as $k => $splitAttribute) {
                foreach ($splitAttribute as $attr => $split) {
                    $xml->writeAttribute($attr, $split);
                }
            }
            $xml->writeAttribute("topLeftCell", $topLeftCell);
            $xml->writeAttribute("activePane", $activePane);
            $xml->writeAttribute("state", "frozen");
            //looks like this: <pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>
            $xml->endElement();
        }
    }


    /**
     * Converts an integer into the alphabet base (A-Z).
     *
     * @param int $n This is the number to convert.
     * @return string The converted number.
     * @author Theriault http://www.php.net/manual/en/function.base-convert.php
     *
     */
    private function _num2alpha($n) {
        $r = '';
        for ($i = 1; $n >= 0 && $i < 10; $i++) {
            $r = chr(0x41 + ($n % pow(26, $i) / pow(26, $i - 1))) . $r;
            $n -= pow(26, $i);
        }
        return $r;
    }

    /**
     * Write the teamplate to disk with the specified filename
     */
    private function _writeTemplate() {
        file_put_contents($this->_filename, base64_decode(self::TEMPLATE));
    }

    /**
     * zip the document
     */
    private function _updateWorksheet() {
        //add it to the template zip
        chdir($this->_tempFolderName);
        echo shell_exec("zip -r $this->_filename xl");
        chdir("./..");
    }

    /**
     * Validate the input. if all is well, run it through the process function. otherwise, return false
     * @param String $infile Input Filename
     * @param Stringe $outfile Output Filename
     * @return boolean Success
     */
    public function writeXLSX($outfile = null) {
        $this->_error = false;
        $this->_exec = false;
        ob_start();
        if (!empty($this->_inputArray) && is_array($this->_inputArray)) {
            $uniq = uniqid();
            $file = $this->_file = tmpfile();

            $dir = str_replace("//", "/", $this->_tempDirectoryBase . "/$uniq");
            $this->_tempFolderName = $dir . "/";
            mkdir($dir);
            if (!$outfile) {
                $file = $dir . '/download.xlsx';
                $outfile = $file;
                $return = true;
            } else {
                $return = false;
            }
            $this->_filename = $outfile;

            //we have a file so do stuff
            try {
                $this->_writeXML();
                $this->_writeTemplate();
                $this->_updateWorksheet();
                $this->_error = false;
            } catch (\Exception $ex) {
                $this->_error = "An error occured. {$ex->getMessage()} on line {$ex->getLine()}";
            }
        } else {
            $this->_error = "You have not loaded a csv or array.";
        }
        $this->_exec = ob_get_clean();
        if (!$return) {
            return !$this->_error;
        } else {
            $retVal = file_get_contents($file);
            unlink($file);
            shell_exec("rm -rf $dir");
            return $retVal;
        }
    }

    /**
     * Serve up an xlsx file. Accepts a filepath or a string as input.
     * @param type $infile
     */
    public function serveXLSX($file, $isFile = false) {
        //clear buffer
        while (ob_get_level()) {
            ob_end_clean();
        }

        //string or file
        /*
          if (file_exists($file)) {
          $file = file_get_contents($file);
          }
         */

        if ($isFile) {
            $len = strlen(file_get_contents($file));
        } else {
            $len = strlen($file);
        }


        //application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
        //serve up the file
        header('Content-type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment; filename="' . "download.xlsx" . '"');
        header('Content-Length: ' . $len);
        header('Content-Transfer-Encoding: binary');
        header('Expires: 0');
        header('Pragma: no-cache');
        if ($isFile) {
            readfile($file);
        } else {
            echo $file;
        }
        exit;
        //delete the file after it has been served up
        //unlink($file);
        //rmdir($dir);
    }

    /**
     * Return the last error
     * @return array|boolean false if no errors
     */
    public function getLastError() {
        if (!$this->error) {
            return false;
        } else {
            return array(
                'errorText' => $this->_error,
                'consoleText' => $this->_exec,
            );
        }
    }

    /**
     * This is where the work is done. Write the array data to XML
     */
    private function _writeXML() {

        $rows = $this->_inputArray;


        //echo var_export($rows, true);

        $colMin = 1;
        $colMax = count($rows[0]);
        $spansAttribute = "$colMin:$colMax";


        //$rows = explode("\r\n", $csv);
        $rowMin = 1;
        $rowMax = count($rows);


//set up the output

        mkdir($this->_tempFolderName . "/xl");
        mkdir($this->_tempFolderName . "/xl/worksheets/");

//set up xml
        $xml = new XMLWriter;
        $xml->openUri($this->_tempFolderName . "/xl/worksheets/sheet1.xml");
        $xml->startDocument('1.0', 'UTF-8');
        $xml->setIndent(true);

//open work sheet
        $xml->startElement("worksheet");
        $xml->writeAttribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        $xml->writeAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

//write dimension tag
        $xml->startElement("dimension");
//calculate the size (reference area). start is A1, end is bottom left corner
        $xml->writeAttribute("ref", "A1:" . $this->_num2alpha($colMax) . $rowMax);
        $xml->endElement();


//make the sheetViews, which is the workbook/sheet structure. we just need 1 sheet, with A1 selected
//start sheetViews
        $xml->startElement("sheetViews");
//start first sheet
        $xml->startElement("sheetView");
//select 1st tab
        $xml->writeAttribute("workbookViewId", "0");
        $xml->writeAttribute("tabSelected", "1");
        if (!empty($this->_freezePanes) && in_array($this->_freezePanes['type'], ["top", "left"])) {
            //this adds the pane, and also makes the selected element the pane
            $this->_addPaneElement($xml);
        } else {
            //make the selected element the first cell
            //select cell A1
//open selection
            $xml->startElement("selection");
            $xml->writeAttribute("activeCell", "A1");
            $xml->writeAttribute("sqref", "A1");
//close selection
            $xml->endElement();
        }
//close sheetView
        $xml->endElement();
//close sheetViews
        $xml->endElement();

//set the row height
        $xml->startElement("sheetFormatPr");
        $xml->writeAttribute("defaultRowHeight", "12.75");
        $xml->endElement();

//set the number of columns and width
        $xml->startElement("cols");
        $xml->startElement("col");
        $xml->writeAttribute("min", $colMin);
        $xml->writeAttribute("max", $colMax);
        $xml->writeAttribute("width", 10);
        $xml->endElement();
        $xml->endElement();


//write the spreadsheet
        $xml->startElement("sheetData");


        foreach ($rows as $r => $row) {
            $xml->startElement("row");
            $xml->writeAttribute("r", $r + 1);
            $xml->writeAttribute("spans", $spansAttribute);
            //$row = explode(",", $row);


            foreach ($row as $c => $col) {

                //start the c(ell)
                $xml->startElement("c");

                //write the row attribute ..eg A5

                $xml->writeAttribute("r", $this->_num2alpha($c) . ($r + 1));

                //write the t attribute if its a string.
                //t attribute: t=s: shared string, t=str: string, no t attribute:number
                if (!empty($col) && !is_numeric($col)) {
                    $xml->writeAttribute("t", "str");
                } //else no t attribute needed
                //finally, write the value if it isn't empty
                if (!empty($col)) {
                    if (is_numeric($col)) {
                        $v = trim($col);
                    } else {
                        $v = $col;
                    }
                    $xml->writeElement("v", $v);
                }
                //end the c(ell)
                $xml->endElement();
            }
            //end the row
            $xml->endElement();
        }
//end the spreadsheet
        $xml->endElement();

//print options, empty element
        $xml->writeElement("printOptions");

//default margins
        $xml->startElement("pageMargins");
        $xml->writeAttribute("left", "1");
        $xml->writeAttribute("right", "1");
        $xml->writeAttribute("top", "1.667");
        $xml->writeAttribute("bottom", "1.667");
        $xml->writeAttribute("header", "1");
        $xml->writeAttribute("footer", "1");
        $xml->endElement();

//page setup
        /*
          $xml->startElement("pageSetup");
          $xml->writeAttribute("orientation", "portrait");
          $xml->writeAttribute("r:id", "rId1");
          $xml->endElement();
         */
//close worksheet
        $xml->endElement();
        $xml->flush();
    }

}