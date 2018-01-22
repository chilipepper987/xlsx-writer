<?php

/**
 * Class XLSXStyleSheet
 * Handles the stylesheet portion (styles.xml) of an xlsx document, as well as the cell assignments for each style
 */
class XLSXStyleSheet {
    private $_font, $_border, $_fill, $_style_xml;

    /**
     * XLSXStyleSheet constructor.
     * @param $xlsx_document_root String path to the root directory of the xlsx document (ie, the one containing the xl, _rels, docProps folders)
     * @throws Exception if styles.xml is not found at $xlsx_document_root/xl/styles.xml
     */
    public function __construct($xlsx_document_root) {
        $xlsx_document_root = pathinfo($xlsx_document_root, PATHINFO_DIRNAME);
        $stylesheet = $xlsx_document_root . '/xl/styles.xml';
        if (!file_exists($stylesheet)) {
            throw new \Exception("No styles.xml file found in $xlsx_document_root/xl/");
        }
        $this->_style_xml = $stylesheet;
    }

    /**
     * Add the <i>font</i> styles specified in array <code>$options</code> to a <code>$target</code> of type <code>$targetType</code>
     * @param String $targetType row|column|cell|range
     * @param String $target rowNumber|columnNumber|cellAddress(eg, B2)|range(A2:B6)
     * @param Array $options An array of font options...
     * @return Boolean success
     */
    public function addFont($targetType, $target, $options) {
        return $this->_addStyle($targetType, $target, $options, "font");
    }

    /**
     * Add the <i>border</i> styles specified in array <code>$options</code> to a <code>$target</code> of type <code>$targetType</code>
     * @param String $targetType row|column|cell|range
     * @param String $target rowNumber|columnNumber|cellAddress(eg, B2)|range(A2:B6)
     * @param Array $options An array of border options...
     * @return Boolean success
     */
    public function addBorder($targetType, $target, $options) {
        return $this->_addStyle($targetType, $target, $options, "border");
    }

    /**
     * Add the <i>fill</i> styles specified in array <code>$options</code> to a <code>$target</code> of type <code>$targetType</code>
     * @param String $targetType row|column|cell|range
     * @param String $target rowNumber|columnNumber|cellAddress(eg, B2)|range(A2:B6)
     * @param Array $options An array of font options...
     * @return Boolean success
     */
    public function addFill($targetType, $target, $options) {
        return $this->_addStyle($targetType, $target, $options, "fill");
    }

    private function _addStyle($targetType, $target, $options, $style) {
        if (!in_array($style, array("font", "border", "fill"))) {
            return false;
        } else {
            $styleArr = $this->{"_" . $style};
        }
        $targetType = strtolower($targetType);
        if (in_array($targetType, ["row", "column"])) {
            $target = intval($target);
        } elseif (in_array($targetType, ["cell", "range"])) {
            //no range support yet
            return false;
        } else {
            return false;
            //invalid target type
        }
        if (!empty($styleArr[$targetType][$target])) {
            $styleArr[$targetType][$target] = $options;
        } else {
            //merge the options. new ones will overwrite old ones
            //for array (+), contents from the left side are preserved, and
            //overwrite contents from the right side, so
            $styleArr[$targetType][$target] = ($options + $styleArr[$targetType][$target]);
        }
        //save it back
        $this->{"_" . $style} = $styleArr;
    }

    protected function _writeStylesheet() {
        //we have a file. write the basic stylesheet.

        //look through options to see if we have any extra styles to write beyond the default
        //font defaults
        ob_start();
        ?>
        <sz val="11"/>
        <color theme="1"/>
        <name val="Calibri"/>
        <family val="2"/>
        <scheme val="minor"/><?php

        $fontDefaults = ob_get_clean();
        //fill defaults
        $defaultFills = array(
            array("patternFill" => array("patternType" => "none")),
            array("patternFill" => array("patternType" => "gray125")),
        );
        //border defaults
        $defaultBorders = array(
            array(
                "left" => false,
                "right" => false,
                "top" => false,
                "bottom" => false,
                "diagonal" => false,
            ),
        );
        //styles, we have cellStyleXfs, cellXfs, and cellStyles. we are concerned with cellXfs


        $xml = new XMLWriter;
        $xml->openUri($this->_style_xml);
        $xml->startDocument('1.0', 'UTF-8');
        $xml->setIndent(true);

        //open stylesheet
        $xml->startElement("styleSheet");
        $xml->writeAttribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        $xml->startElement("fonts");
        $xml->writeAttribute("count", count($this->_fonts));
        foreach ($fonts as $font) {
            $xml->startElement("font");
            if ($font['style']) {
                $xml->writeElement($fontStyle);
            }
            //font defaults
            $xml->writeRaw($fontDefaults);
            //end font
            $xml->endElement();
        }
        //end fonts
        $xml->endElement();
    }
}