<?php
namespace Rajwebsoft\Xlsreader\Library;
use Exception;

class ITFXlsReader
{
    const DS = DIRECTORY_SEPARATOR;
    const Date_Bias = 25569; // number of days between Excel and UNIX epoch
    const VERSION = "1.6";
    const TPL_DIR = "templates";
    private $_parsed = [];
    private $arrXMLs = []; // all XML files
    private $arrSheets = []; // all sheets
    private $arrSheetPath = []; // all paths to sheets
    private $_cSheet; // current sheet
    /** @ignore */
    public $defaultFileName = "ITFXlsReader.xlsx";
    /** @ignore */
    static $arrIndexedColors = [
        "00000000",
        "00FFFFFF",
        "00FF0000",
        "0000FF00",
        "000000FF",
        "00FFFF00",
        "00FF00FF",
        "0000FFFF",
        "00000000",
        "00FFFFFF",
        "00FF0000",
        "0000FF00",
        "000000FF",
        "00FFFF00",
        "00FF00FF",
        "0000FFFF",
        "00800000",
        "00008000",
        "00000080",
        "00808000",
        "00800080",
        "00008080",
        "00C0C0C0",
        "00808080",
        "009999FF",
        "00993366",
        "00FFFFCC",
        "00CCFFFF",
        "00660066",
        "00FF8080",
        "000066CC",
        "00CCCCFF",
        "00000080",
        "00FF00FF",
        "00FFFF00",
        "0000FFFF",
        "00800080",
        "00800000",
        "00008080",
        "000000FF",
        "0000CCFF",
        "00CCFFFF",
        "00CCFFCC",
        "00FFFF99",
        "0099CCFF",
        "00FF99CC",
        "00CC99FF",
        "00FFCC99",
        "003366FF",
        "0033CCCC",
        "0099CC00",
        "00FFCC00",
        "00FF9900",
        "00FF6600",
        "00666699",
        "00969696",
        "00003366",
        "00339966",
        "00003300",
        "00333300",
        "00993300",
        "00993366",
        "00333399",
        "00333333",
    ];
    public function __construct($templatePath = "")
    {
        if (!$templatePath) {
            $templatePath = "empty";
        } else {
            $this->defaultFileName = basename($templatePath);
        }
        $templatePath = file_exists($templatePath)
            ? $templatePath
            : dirname(__FILE__) .
                self::DS .
                self::TPL_DIR .
                self::DS .
                $templatePath;

        if (!file_exists($templatePath)) {
            throw new ITFXlsReader_Exception(
                "XLSX template ({$templatePath}) does not exist"
            );
        }

        if (is_dir($templatePath)) {
            $ITFXlsReader_FS = new ITFXlsReader_FS($templatePath);
            list($arrDir, $arrFiles) = $ITFXlsReader_FS->get();
        } else {
            list($arrDir, $arrFiles) = $this->unzipToMemory($templatePath);
        }

        $nSheets = 0;
        $nFirstSheet = 1;
        foreach ($arrFiles as $path => $contents) {
            $path = "/" . str_replace(self::DS, "/", $path);

            $this->arrXMLs[$path] = @simplexml_load_string($contents);

            if (empty($this->arrXMLs[$path])) {
                $this->arrXMLs[$path] = (string) $contents;
            }

            if (preg_match("/\.rels$/", $path)) {
                foreach ($this->arrXMLs[$path]->Relationship as $Relationship) {
                    if (
                        (string) $Relationship["Type"] ==
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
                    ) {
                        $this->officeDocumentPath = self::getPathByRelTarget(
                            $path,
                            (string) $Relationship["Target"]
                        );
                        $this->officeDocument =
                            &$this->arrXMLs[$this->officeDocumentPath];
                    }

                    if (
                        (string) $Relationship["Type"] ==
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
                    ) {
                        $this->sharedStrings =
                            &$this->arrXMLs[
                                self::getPathByRelTarget(
                                    $path,
                                    (string) $Relationship["Target"]
                                )
                            ];
                    }

                    if (
                        (string) $Relationship["Type"] ==
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
                    ) {
                        $this->theme =
                            &$this->arrXMLs[
                                self::getPathByRelTarget(
                                    $path,
                                    (string) $Relationship["Target"]
                                )
                            ];
                    }

                    if (
                        (string) $Relationship["Type"] ==
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
                    ) {
                        $this->styles =
                            &$this->arrXMLs[
                                self::getPathByRelTarget(
                                    $path,
                                    (string) $Relationship["Target"]
                                )
                            ];
                    }
                }
            }
        }

        $this->officeDocumentRelPath = self::getRelFilePath(
            $this->officeDocumentPath
        );
        $ix = 0;
        foreach ($this->officeDocument->sheets->sheet as $sheet) {
            $relId = $sheet->attributes("r", true)->id;

            foreach (
                $this->arrXMLs[$this->officeDocumentRelPath]->Relationship
                as $Relationship
            ) {
                if ((string) $Relationship["Id"] == $relId) {
                    $path = self::getPathByRelTarget(
                        $this->officeDocumentRelPath,
                        (string) $Relationship["Target"]
                    );
                    break;
                }
            }
            $this->arrSheets[(string) $sheet["sheetId"]] =
                &$this->arrXMLs[$path];
            $this->arrSheetPath[(string) $sheet["sheetId"]] = $path;
            $nFirstSheet = $ix == 0 ? (string) $sheet["sheetId"] : $nFirstSheet;
            $ix++;
        }

        $this->selectSheet($nFirstSheet);
    }

    public function data($cellAddress, $data = null, $t = "s")
    {
        $retVal = null;

        list($x, $y, $addrA1, $addrR1C1) = self::cellAddress($cellAddress);

        $c = $this->locateCell($x, $y);
        if (!$c && $data !== null) {
            $c = &$this->addCell($x, $y);
        }

        if (isset($c->v[0])) {
            // if it has value

            $o_v = &$c->v[0];
            if ($c["t"] == "s") {
                // if existing type is string
                $siIndex = (int) $c->v[0];
                $o_si = &$this->sharedStrings->si[$siIndex];
                $retVal = strip_tags($o_si->asXML()); //return plain string without formatting
            } else {
                // if not or undefined
                $retVal = $this->formatDataRead($c["s"], (string) $o_v);
                if ($data !== null && $t == "s") {
                    // if forthcoming type is string, we add shared string
                    $o_si = &$this->addSharedString($c);
                }
            }
        } else {
            $retVal = null;
            if (
                $data !== null &&
                !(!is_object($data) && (string) $data == "")
            ) {
                if ($t == "s") {
                    // if we'd like to set data and not to empty this cell
                    // if forthcoming type is string, we add shared string
                    $o_si = &$this->addSharedString($c);
                    $o_v = &$c->v[0];
                } else {
                    // if not, value is inside '<v>' tag
                    $c->addChild("v", $data);
                    $o_v = &$c->v[0];
                }
            }
        }

        if ($data !== null) {
            // if we set data

            if (!is_object($data) && (string) $data == "") {
                // if there's an empty string, we demolite existing data
                unset($c["t"]);
                unset($c->v[0]);
            } else {
                // we set received value
                unset($c->f[0]); // remove forumla
                if (is_numeric($data) && func_num_args() == 2) {
                    // if default
                    $t = "n";
                }
                switch ($t) {
                    case "s":
                        $this->updateSharedString($o_si, $data);
                        break;
                    default:
                        $this->formatDataWrite($t, $data, $c);
                        break;
                }
            }
        }

        return $retVal;
    }

    public function getDataValidationList($cellAddress)
    {
        if ($this->_cSheet->dataValidations->dataValidation) {
            foreach (
                $this->_cSheet->dataValidations->dataValidation
                as $ix => $val
            ) {
                if ($val["type"] != "list") {
                    continue;
                }
                $range = $val["sqref"];
                if (self::checkAddressInRange($cellAddress, $range)) {
                    $ref = (string) $val->formula1[0];
                    break;
                }
            };
        }

        if (!$ref) {
            foreach ($this->_cSheet->extLst->ext as $ext) {
                if ($ext["uri"] != "{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}") {
                    continue;
                }
                $arrNS = $ext->getNamespaces(true);
                foreach ($arrNS as $prfx => $uri) {
                    if (preg_match('/^x[0-9]*$/', $prfx)) {
                        $nsX = $uri;
                        break;
                    }
                }

                $chdn = $ext->children($nsX);

                foreach (
                    $chdn->dataValidations->dataValidation
                    as $ix => $val
                ) {
                    $range = (string) $val->children("xm", true);
                    if (self::checkAddressInRange($cellAddress, $range)) {
                        $ref = (string) $val->formula1[0]->children("xm", true);
                        break;
                    }
                }

                if ($ref) {
                    break;
                }
            };
        }

        return $ref ? $this->getDataByRange($ref) : null;
    }

    public function getDataByRange($range)
    {
        $arrRet = [];
        $diffSheetName = $refSheetID = "";

        $range = preg_replace('/\$([a-z0-9]+)/i', '$1', $range);

        $arrRef = explode("!", $range);

        $range = $arrRef[count($arrRef) - 1];

        if ($diffSheetName = count($arrRef) > 1 ? $arrRef[0] : "") {
            if (!($refSheetID = $this->findSheetByName($diffSheetName))) {
                return false;
            }

            foreach ($this->arrSheets as $id => $sheet) {
                if ($sheet === $this->_cSheet) {
                    $curSheetID = $id;
                    break;
                }
            }

            $this->selectSheet($refSheetID);
        }

        try {
            list($aX, $aY) = self::getRangeArea($range);
        } catch (ITFXlsReader_Exception $e) {
            return false;
        }

        for ($x = $aX[0]; $x <= $aX[1]; $x++) {
            for ($y = $aY[0]; $y <= $aY[1]; $y++) {
                $addr = "R{$y}C{$x}";
                $dt = $this->data($addr);
                if ($dt) {
                    $arrRet[$addr] = $dt;
                }
            };
        }

        if ($diffSheetName) {
            $this->selectSheet($curSheetID);
        }

        return $arrRet;
    }

    public static function checkAddressInRange($adrNeedle, $adrHaystack)
    {
        list($xNeedle, $yNeedle) = self::cellAddress($adrNeedle);

        $arrHaystack = explode(" ", $adrHaystack);
        foreach ($arrHaystack as $range) {
            list($x, $y) = self::getRangeArea($range);

            if (
                $x[0] <= $xNeedle &&
                $xNeedle <= $x[1] &&
                $y[0] <= $yNeedle &&
                $yNeedle <= $y[1]
            ) {
                return true;
            }
        }

        return false;
    }

    public static function getRangeArea($range)
    {
        $arrRng = explode(":", $range);

        list($x[0], $y[0]) = self::cellAddress($arrRng[0]);

        if ($arrRng[1]) {
            list($x[1], $y[1]) = self::cellAddress($arrRng[1]);
        } else {
            $x[1] = $x[0];
            $y[1] = $y[0];
        }

        sort($x);
        sort($y);

        return [$x, $y];
    }

    public function getRowCount()
    {
        $lastRowIndex = 1;
        foreach ($this->_cSheet->sheetData->row as $row) {
            $lastRowIndex = (int) $row["r"];
        }
        return $lastRowIndex;
    }

    public function fill($cellAddress, $fillColor)
    {
        $fillColor = $fillColor ? self::colorW3C2Excel($fillColor) : "";
        list($x, $y, $addrA1, $addrR1C1) = self::cellAddress($cellAddress);
        $c = &$this->locateCell($x, $y);

        if ($c === null) {
            throw new ITFXlsReader_Exception(
                "cannot apply fill - no cell at " . $cellAddress
            );
        }

        if ($fillColor) {
            $ix = 0;
            foreach ($this->styles->fills->fill as $fill) {
                if (
                    strtoupper((string) $fill->patternFill->fgColor["rgb"]) ==
                    $fillColor
                ) {
                    $fillIx = $ix;
                    break;
                }
                $ix++;
            }
            if (!isset($fillIx)) {
                $xmlFill = simplexml_load_string(
                    "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"{$fillColor}\"/><bgColor indexed=\"64\"/></patternFill></fill>"
                );
                $this->insertElementByPosition(
                    (int) $this->styles->fills["count"],
                    $xmlFill,
                    $this->styles->fills
                );
                $fillIx = (int) $this->styles->fills["count"];
                $this->styles->fills["count"] =
                    (int) $this->styles->fills["count"] + 1;
            }
        } else {
            $fillIx = 0;
        } //http://openxmldeveloper.org/discussions/formats/f/14/p/716/3685.aspx :

        if ($c["s"]) {
            $cellXf = $this->styles->cellXfs->xf[(int) $c["s"]];
            if ((int) $cellXf["fillId"] != $fillIx) {
                // if style is getting changed, we try to locate changed one, if we fail ,we add
                $ix = 0;
                foreach ($this->styles->cellXfs->xf as $xf) {
                    if (
                        (string) $xf["borderId"] ==
                            (string) $cellXf["borderId"] &&
                        (string) $xf["fillId"] == $fillIx &&
                        (string) $xf["fontId"] == (string) $cellXf["fontId"] &&
                        (string) $xf["numFmtId"] ==
                            (string) $cellXf["numFmtId"] &&
                        (string) $xf["xfId"] == (string) $cellXf["xfId"] &&
                        (string) $xf["applyFill"] ==
                            (string) $cellXf["applyFill"]
                    ) {
                        $styleIx = $ix;
                        break;
                    }
                    $ix++;
                }
                if (isset($styleIx)) {
                    $c["s"] = $styleIx;
                } else {
                    $xmlXF = simplexml_load_string(
                        $this->styles->cellXfs->xf[(int) $c["s"]]->asXML()
                    );
                    $xmlXF["fillId"] = $fillIx;
                    $xmlXF["applyFill"] = "1";
                    $this->insertElementByPosition(
                        (int) $this->styles->cellXfs["count"],
                        $xmlXF,
                        $this->styles->cellXfs
                    );
                    $styleIx = (int) $this->styles->cellXfs["count"];
                    $this->styles->cellXfs["count"] =
                        (int) $this->styles->cellXfs["count"] + 1;
                    $c["s"] = $styleIx; // update cell with style
                }
            }
        } else {
            if ($fillIx !== 0) {
                $xmlXF = simplexml_load_string(
                    "<xf borderId=\"0\" fillId=\"{$fillIx}\" fontId=\"0\" numFmtId=\"0\" xfId=\"0\" applyFill=\"1\"/>"
                );
                $this->insertElementByPosition(
                    (int) $this->styles->cellXfs["count"],
                    $xmlXF,
                    $this->styles->cellXfs
                );
                $styleIx = (int) $this->styles->cellXfs["count"];
                $this->styles->cellXfs["count"] =
                    (int) $this->styles->cellXfs["count"] + 1;
                $c["s"] = $styleIx; // update cell with style
            }
        }

        return $c;
    }

    public function getFillColor($cellAddress)
    {
        list($x, $y, $addrA1, $addrR1C1) = self::cellAddress($cellAddress);
        $c = &$this->locateCell($x, $y);

        if ($c === null) {
            throw new ITFXlsReader_Exception(
                "cannot apply fill - no sheet at " . $cellAddress
            );
        }
        if ($c["s"]) {
            $cellXf = $this->styles->cellXfs->xf[(int) $c["s"]];
            $fillIx = (int) $cellXf["fillId"];
            $fgColor =
                $this->styles->fills->fill[$fillIx]->patternFill->fgColor;
            if ($fgColor["rgb"]) {
                return $fgColor["rgb"];
            } else {
                if ($fgColor["theme"]) {
                    return $this->getThemeColor($fgColor["theme"]);
                } elseif ($fgColor["indexed"]) {
                    return self::colorExcel2W3C(
                        self::$arrIndexedColors[(int) $fgColor["indexed"]]
                    );
                }
            }
            $color =
                $this->styles->fills->fill[$fillIx]->patternFill->fgColor[
                    "rgb"
                ];
            if ($color) {
                return $color;
            } else {
                return "#FFFFFF";
            }
        } else {
            return "#FFFFFF";
        }

        return $c;
    }

    protected function getThemeColor($theme)
    {
        $ixScheme = 0;
        foreach (
            $this->theme->children("a", true)->themeElements[0]->clrScheme[0]
            as $ix => $scheme
        ) {
            if ((int) $theme == $ixScheme) {
                foreach ($scheme as $node => $chl) {
                    $domch = dom_import_simplexml($chl);
                    switch ($node) {
                        case "srgbClr":
                        default:
                            return "#" . $domch->getAttribute("val");
                        case "sysClr":
                            return "#" . $domch->getAttribute("lastClr");
                    }
                }
                break;
            }
            $ixScheme++;
        }
    }
    public function cloneRow($ySrc, $yDest)
    {
        $oSrc = $this->locateRow($ySrc);
        if (!$oSrc) {
            return null;
        }

        $domSrc = dom_import_simplexml($oSrc);
        $oDest = simplexml_import_dom($domSrc->cloneNode(true));
        foreach ($oDest->c as $c) {
            unset($c["t"]);
            unset($c->v[0]);
            list($x) = self::cellAddress($c["r"]);
            if (preg_match("/^R([0-9]+)C([0-9]+)$/i", $c["r"])) {
                $c["r"] = "R{$yDest}C{$x}";
            } else {
                $c["r"] = $this->index2letter($x) . "{$yDest}";
            }
        }

        $oDest["r"] = $yDest;

        $retVal = $this->insertElementByPosition(
            $yDest,
            $oDest,
            $this->_cSheet->sheetData
        );

        $this->shiftDownMergedCells($yDest, $ySrc);

        return $retVal;
    }
    public function findSheetByName($name)
    {
        foreach ($this->officeDocument->sheets->sheet as $sheet) {
            if ((string) $sheet["name"] == $name) {
                return (string) $sheet["sheetId"];
            }
        }

        return false;
    }
    public function selectSheet($id)
    {
        if (!isset($this->arrSheets[$id])) {
            throw new ITFXlsReader_Exception('can\'t select sheet #' . $id);
        }
        $this->_cSheet = $this->arrSheets[$id];
        return $this;
    }
    public function cloneSheet($originSheetId, $newSheetName = "")
    {
        if (!isset($this->arrSheets[$originSheetId])) {
            throw new ITFXlsReader_Exception(
                'can\'t select sheet #' . $originSheetId
            );
        }
        $maxID = 1;
        $maxSheetFileIX = 1;
        foreach ($this->officeDocument->sheets->sheet as $sheet) {
            $maxID = max((int) $sheet["sheetId"], $maxID);
            $sheetFileName = basename(
                $this->arrSheetPath[(string) $sheet["sheetId"]]
            );
            preg_match("/sheet([0-9]+)\.xml/", $sheetFileName, $arrMatch);
            $sheetFileIX = (int) $arrMatch[1];
            $maxSheetFileIX = max($maxSheetFileIX, $sheetFileIX);
            $relId = $sheet->attributes("r", true)->id;
            $maxRelID = max($maxRelID, (int) str_replace("rId", "", $relId));
        }
        $newSheetID = $maxID + 1;
        $newSheetRelID = "rId" . ($maxRelID + 1);
        $newSheetFileName = "sheet" . ($maxSheetFileIX + 1) . ".xml";
        $newSheetFullPath =
            dirname($this->arrSheetPath[(string) $originSheetId]) .
            "/" .
            $newSheetFileName;
        $newSheetName = $newSheetName ? $newSheetName : "Sheet {$newSheetID}";
        $this->arrXMLs[$newSheetFullPath] = simplexml_load_string(
            $this->arrSheets[(string) $originSheetId]->asXML()
        );
        if (
            isset(
                $this->arrXMLs[
                    self::getRelFilePath($this->arrSheetPath[$originSheetId])
                ]
            )
        ) {
            $this->arrXMLs[
                self::getRelFilePath($newSheetFullPath)
            ] = simplexml_load_string(
                $this->arrXMLs[
                    self::getRelFilePath($this->arrSheetPath[$originSheetId])
                ]->asXML()
            );
        }
        $newSh = $this->officeDocument->sheets->addChild("sheet");
        $newSh->addAttribute(
            "r:id",
            $newSheetRelID,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        );
        $newSh->addAttribute("sheetId", $newSheetID);
        $newSh->addAttribute("name", $newSheetName);
        // <Relationship Target="worksheets/sheet5.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Id="rId5"/>
        $newRel = $this->arrXMLs[$this->officeDocumentRelPath]->addChild(
            "Relationship"
        );
        $newRel->addAttribute("Target", "worksheets/" . $newSheetFileName);
        $newRel->addAttribute(
            "Type",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
        );
        $newRel->addAttribute("Id", $newSheetRelID);
        //<Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" PartName="/xl/worksheets/sheet5.xml"/>
        $newOvr = $this->arrXMLs["/[Content_Types].xml"]->addChild("Override");
        $newOvr->addAttribute(
            "ContentType",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
        );
        $newOvr->addAttribute(
            "PartName",
            "/xl/worksheets/" . $newSheetFileName
        );
        $this->updateWorkbookLinks();

        return (string) $newSheetID;
    }
    public function renameSheet($sheetId, $newName)
    {
        if (!isset($this->arrSheets[$sheetId])) {
            throw new ITFXlsReader_Exception('can\'t get sheet #' . $sheetId);
        }

        foreach ($this->officeDocument->sheets->sheet as $sheet) {
            if ((string) $sheet["sheetId"] == (string) $sheetId) {
                $sheet["name"] = $newName;
                break;
            }
        }
        $this->updateAppXML();
    }
    public function removeSheet($id)
    {
        $sheetXMLFileName = $this->arrSheetPath[(string) $id];
        $sheetXMLRelsFileName = self::getRelFilePath($sheetXMLFileName);
        if ($this->arrXMLs[$sheetXMLRelsFileName]->Relationship) {
            foreach (
                $this->arrXMLs[$sheetXMLRelsFileName]->Relationship
                as $Relationship
            ) {
                unset(
                    $this->arrXMLs[
                        self::getPathByRelTarget(
                            $sheetXMLRelsFileName,
                            $Relationship["Target"]
                        )
                    ]
                );
            };
        }
        unset($this->arrXMLs[$sheetXMLRelsFileName]);
        unset($this->arrXMLs[$sheetXMLFileName]);
        unset($this->arrSheets[(string) $id]);
        unset($this->arrSheetPath[(string) $id]);
        $ix = 0;
        foreach ($this->officeDocument->sheets->sheet as $sheet) {
            $relId = $sheet->attributes("r", true)->id; // take old relId

            $ixRel = 0;
            foreach (
                $this->arrXMLs[$this->officeDocumentRelPath]->Relationship
                as $Relationship
            ) {
                if ((string) $Relationship["Id"] == $relId) {
                    break;
                }
                $ixRel++;
            }

            if ((string) $sheet["sheetId"] == (string) $id) {
                unset(
                    $this->arrXMLs[$this->officeDocumentRelPath]->Relationship[
                        $ixRel
                    ]
                );
                $ixToDel = $ix;
                break;
            }

            $ix++;
        }
        unset($this->officeDocument->sheets->sheet[$ixToDel]);
        $ixDel = $nCount = 0;
        foreach (
            $this->arrXMLs["/[Content_Types].xml"]->Override
            as $Override
        ) {
            if ((string) $Override["PartName"] == $sheetXMLFileName) {
                $ixDel = $nCount;
            }
            $nCount++;
        }
        unset($this->arrXMLs["/[Content_Types].xml"]->Override[$ixDel]);

        $this->updateWorkbookLinks();
    }
    /**
     * @ignore
     */
    protected function getPathByRelTarget($relFilePath, $targetPath)
    {
        $relFileDirectory = preg_replace(
            "/(_rels)$/",
            "",
            dirname($relFilePath)
        );
        $arrPath = explode("/", rtrim($relFileDirectory, "/"));
        $arrTargetPath = explode("/", ltrim($targetPath, "/"));
        foreach ($arrTargetPath as $directory) {
            switch ($directory) {
                case ".":
                    break;
                case "..":
                    if (isset($arrPath[count($arrPath) - 1])) {
                        unset($arrPath[count($arrPath) - 1]);
                    } else {
                        throw new Exception(
                            "Unable to change directory upwards (..)"
                        );
                    }
                    break;
                default:
                    $arrPath[] = $directory;
                    break;
            }
        }

        return implode("/", $arrPath);
    }
    protected function getRelFilePath($xmlPath)
    {
        return dirname($xmlPath) .
            "/_rels" .
            str_replace(dirname($xmlPath), "", $xmlPath) .
            ".rels";
    }
    /** @ignore  */
    private function updateSharedString($o_si, $data)
    {
        $dom_si = dom_import_simplexml($o_si);

        while ($dom_si->hasChildNodes()) {
            $dom_si->removeChild($dom_si->firstChild);
        }

        if (!is_object($data)) {
            $data = simplexml_load_string(
                "<richText><t>" . htmlspecialchars($data) . "</t></richText>"
            );
        }

        foreach ($data->children() as $childNode) {
            $domInsert = $dom_si->ownerDocument->importNode(
                dom_import_simplexml($childNode),
                true
            );
            $dom_si->appendChild($domInsert);
        }

        return simplexml_import_dom($o_si);
    }
    /** @ignore  */
    private function formatDataRead($style, $data)
    {
        if ((string) $style == "") {
            return (string) $data;
        }

        $numFmt = (string) $this->styles->cellXfs->xf[(int) $style]["numFmtId"];

        switch ($numFmt) {
            case "14": // = 'mm-dd-yy';
            case "15": // = 'd-mmm-yy';
            case "16": // = 'd-mmm';
            case "17": // = 'mmm-yy';
            case "18": // = 'h:mm AM/PM';
            case "19": // = 'h:mm:ss AM/PM';
            case "20": // = 'h:mm';
            case "21": // = 'h:mm:ss';
            case "22": // = 'm/d/yy h:mm';
                return date("Y-m-d", 60 * 60 * 24 * ($data - self::Date_Bias));
            default:
                if ((int) $numFmt >= 164) {
                    //look for custom format number
                    foreach ($this->styles->numFmts[0]->numFmt as $o_numFmt) {
                        if ((int) $o_numFmt["numFmtId"] == (int) $numFmt) {
                            $formatCode = (string) $o_numFmt["formatCode"];
                            if (preg_match("/[dmyh]+/i", $formatCode)) {
                                // CHECK THIS OUT!!! it's just a guess!
                                return date(
                                    "Y-m-d",
                                    60 * 60 * 24 * ($data - self::Date_Bias)
                                );
                            }
                            break;
                        }
                    }
                }
                return $data;
                break;
        }
    }
    private function addSharedString(&$oCell)
    {
        $ssIndex = count($this->sharedStrings->si);

        $oSharedString = $this->sharedStrings->addChild("si", "");
        $this->sharedStrings["uniqueCount"] = $ssIndex + 1;
        $this->sharedStrings["count"] = $this->sharedStrings["count"] + 1;

        $oCell["t"] = "s";
        if (isset($oCell->v[0])) {
            $oCell->v[0] = $ssIndex;
        } else {
            $oCell->addChild("v", $ssIndex);
        }

        return $oSharedString;
    }
    private function convertDateTime($date_input)
    {
        $days = 0; # Number of days since epoch
        $seconds = 0; # Time expressed as fraction of 24h hours in seconds
        $year = $month = $day = 0;
        $hour = $min = $sec = 0;
        $date_time = $date_input;
        if (preg_match("/(\d{4})\-(\d{2})\-(\d{2})/", $date_time, $matches)) {
            list($junk, $year, $month, $day) = $matches;
        }
        if (preg_match("/(\d{2}):(\d{2}):(\d{2})/", $date_time, $matches)) {
            list($junk, $hour, $min, $sec) = $matches;
            $seconds = ($hour * 60 * 60 + $min * 60 + $sec) / (24 * 60 * 60);
        }
        # Special cases for Excel.
        if ("$year-$month-$day" == "1899-12-31") {
            return $seconds;
        } # Excel 1900 epoch
        if ("$year-$month-$day" == "1900-01-00") {
            return $seconds;
        } # Excel 1900 epoch
        if ("$year-$month-$day" == "1900-02-29") {
            return 60 + $seconds;
        } # Excel false leapday
        # We calculate the date by calculating the number of days since the epoch
        # and adjust for the number of leap days. We calculate the number of leap
        # days by normalising the year in relation to the epoch. Thus the year 2000
        # becomes 100 for 4 and 100 year leapdays and 400 for 400 year leapdays.
        $epoch = 1900;
        $offset = 0;
        $norm = 300;
        $range = $year - $epoch;

        # Set month days and check for leap year.
        $leap = $year % 400 == 0 || ($year % 4 == 0 && $year % 100) ? 1 : 0;
        $mdays = [31, $leap ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];

        # Some boundary checks
        if ($year < $epoch || $year > 9999) {
            return $seconds;
        }
        if ($month < 1 || $month > 12) {
            return $seconds;
        }
        if ($day < 1 || $day > $mdays[$month - 1]) {
            return $seconds;
        }

        # Accumulate the number of days since the epoch.
        $days = $day; # Add days for current month
        $days += array_sum(array_slice($mdays, 0, $month - 1)); # Add days for past months
        $days += $range * 365; # Add days for past years
        $days += intval($range / 4); # Add leapdays
        $days -= intval(($range + $offset) / 100); # Subtract 100 year leapdays
        $days += intval(($range + $offset + $norm) / 400); # Add 400 year leapdays
        $days -= $leap; # Already counted above
        # Adjust for Excel erroneously treating 1900 as a leap year.
        if ($days > 59) {
            $days++;
        }
        return $days + $seconds;
    }
    private function formatDataWrite($type, $data, $c)
    {
        if (isset($c["t"])) {
            unset($c["t"]);
        }

        switch ($type) {
            case "d":
                $c->v[0] = $this->convertDateTime($data);
                break;
            default:
                $c->v[0] = (string) $data;
                break;
        }
    }
    private function locateCell($x, $y)
    {
        $addrA1 = $this->index2letter($x) . $y;
        $addrR1C1 = "R{$y}C{$x}";

        $row = $this->locateRow($y);
        if ($row === null) {
            return null;
        }

        foreach ($row->c as $ixC => $c) {
            if ($c["r"] == $addrA1 || $c["r"] == $addrR1C1) {
                return $c;
            }
        }

        return null;
    }
    private function addCell($x, $y)
    {
        $oValue = null;

        $oRow = $this->locateRow($y);

        if (!$oRow) {
            $oRow = $this->addRow(
                $y,
                simplexml_load_string("<row r=\"{$y}\"></row>")
            );
        }

        $xmlCell = simplexml_load_string(
            "<c r=\"" . $this->index2letter($x) . $y . "\"></c>"
        );
        $oCell = &$this->insertElementByPosition($x, $xmlCell, $oRow);

        return $oCell;
    }
    private function locateRow($y)
    {
        foreach ($this->_cSheet->sheetData->row as $ixRow => $row) {
            if ($row["r"] == $y) {
                return $row;
            }
        }
        return null;
    }
    private function addRow($y, $oRow)
    {
        $this->shiftDownMergedCells($y);

        return $this->insertElementByPosition(
            $y,
            $oRow,
            $this->_cSheet->sheetData
        );
    }
    private function shiftDownMergedCells($yStart, $yOrigin = null)
    {
        if (count($this->_cSheet->mergeCells->mergeCell) == 0) {
            return;
        }

        $toAdd = [];

        foreach ($this->_cSheet->mergeCells->mergeCell as $mergeCell) {
            list($cell1, $cell2) = explode(":", $mergeCell["ref"]);

            list($x1, $y1) = self::cellAddress($cell1);
            list($x2, $y2) = self::cellAddress($cell2);

            if (max($y1, $y2) >= $yStart && min($y1, $y2) < $yStart) {
                // if mergeCells are crossing inserted row
                throw new ITFXlsReader_Exception(
                    "mergeCell {$mergeCell["ref"]} is crossing newly inserted row at {$yStart}"
                );
            }

            if (min($y1, $y2) >= $yStart) {
                $mergeCell["ref"] =
                    $this->index2letter($x1) .
                    ($y1 + 1) .
                    ":" .
                    $this->index2letter($x2) .
                    ($y2 + 1);
            }

            if ($yOrigin !== null) {
                if ($y1 == $y2 && $y1 == $yOrigin) {
                    // if there're merged cells on cloned row we add new <mergeCell>
                    $toAdd[] =
                        $this->index2letter($x1) .
                        $yStart .
                        ":" .
                        $this->index2letter($x2) .
                        $yStart;
                }
            }
        }

        foreach ($toAdd as $newMergeCellRange) {
            $newMC = $this->_cSheet->mergeCells->addChild("mergeCell");
            $newMC["ref"] = $newMergeCellRange;
            $this->_cSheet->mergeCells["count"] =
                $this->_cSheet->mergeCells["count"] + 1;
        }
    }
    private function insertElementByPosition($position, $oInsert, $oParent)
    {
        $domParent = dom_import_simplexml($oParent);
        $domInsert = $domParent->ownerDocument->importNode(
            dom_import_simplexml($oInsert),
            true
        );

        $insertBeforeElement = null;
        $ix = 0;

        foreach ($domParent->childNodes as $element) {
            $el_position = $this->getElementPosition($element, $ix);

            if ($position < $el_position) {
                // if needed element is ahead of current one
                $insertBeforeElement = &$element;
                break;
            }
            if (
                $element->nextSibling !== null &&
                $position <=
                    $this->getElementPosition($element->nextSibling, $ix + 1)
            ) {
                $insertBeforeElement = &$element->nextSibling;
                break;
            }
            $ix++;
        }

        $ix = 0;
        if ($domInsert->nodeName == "row") {
            foreach ($domParent->childNodes as $element) {
                $el_position = $this->getElementPosition($element, $ix);
                if ($el_position >= $position) {
                    $oElement = simplexml_import_dom($element);
                    $oElement["r"] = $el_position + 1; //row 'r' attribute
                    foreach ($oElement->c as $c) {
                        // cells inside it
                        list($x, $y, $a1, $r1c1) = self::cellAddress($c["r"]);
                        $c["r"] =
                            $c["r"] == $a1
                                ? self::index2letter($x) . ($el_position + 1)
                                : "R" . ($el_position + 1) . "C{$x}";
                    }
                }
                $ix++;
            };
        }

        if ($insertBeforeElement !== null) {
            return simplexml_import_dom(
                $domParent->insertBefore($domInsert, $insertBeforeElement)
            );
        } else {
            return simplexml_import_dom($domParent->appendChild($domInsert));
        }
    }
    private function getElementPosition($domXLSXElement, $ix)
    {
        if (count($domXLSXElement->attributes) != 0) {
            foreach ($domXLSXElement->attributes as $ix => $attr) {
                if ($attr->name == "r") {
                    $strPos = (string) $attr->value;
                }
            };
        }

        switch ($domXLSXElement->nodeName) {
            case "row":
                return (int) $strPos;
            case "c":
                list($x) = self::cellAddress($strPos);
                return (int) $x;
            default:
                return $ix;
        }
    }
    private function getRow($y)
    {
        $oRow = null;
        foreach ($this->_cSheet->sheetData->row as $ixRow => $row) {
            if ($row["r"] == $y) {
                $oRow = &$row;
                break;
            }
        }

        if ($oRow === null) {
            $oRow = $this->addRow($y);
        }

        return $oRow;
    }
    public static function cellAddress($cellAddress)
    {
        if (preg_match("/^R([0-9]+)C([0-9]+)$/i", $cellAddress, $arrMatch)) {
            //R1C1 style
            return [
                $arrMatch[2],
                $arrMatch[1],
                self::index2letter($arrMatch[2]) . $arrMatch[1],
                $cellAddress,
            ];
        } else {
            if (preg_match("/^([a-z]+)([0-9]+)$/i", $cellAddress, $arrMatch)) {
                $x = self::letter2index($arrMatch[1]);
                $y = $arrMatch[2];
                return [$x, $y, $cellAddress, "R{$y}C{$x}"];
            }
        }

        throw new ITFXlsReader_Exception(
            "invalid cell address: {$cellAddress}"
        );
    }
    private static function index2letter($index)
    {
        $nLength = ord("Z") - ord("A") + 1;
        $strLetter = "";
        while ($index > 0) {
            $rem = $index % $nLength == 0 ? $nLength : $index % $nLength;
            $strLetter = chr(ord("A") + $rem - 1) . $strLetter;
            $index =
                floor($index / $nLength) - ($index % $nLength == 0 ? 1 : 0);
        }

        return $strLetter;
    }
    private static function colorW3C2Excel($color)
    {
        if (!preg_match("/#[0-9A-F]{2}[0-9A-F]{2}[0-9A-F]{2}/i", $color)) {
            throw new ITFXlsReader_Exception("bad W3C color format: {$color}");
        }
        return strtoupper(preg_replace("/^(#)/", "FF", $color));
    }
    private static function colorExcel2W3C($color)
    {
        if (
            !preg_match(
                "/[0-9A-F]{2}[0-9A-F]{2}[0-9A-F]{2}[0-9A-F]{2}/i",
                $color
            )
        ) {
            throw new ITFXlsReader_Exception(
                "bad OpenXML color format: {$color}"
            );
        }
        return strtoupper(preg_replace("/^([0-9A-F]{2})/i", "#", $color));
    }
    private static function letter2index($strLetter)
    {
        $x = 0;
        $nLength = ord("Z") - ord("A") + 1;
        for ($i = strlen($strLetter) - 1; $i >= 0; $i--) {
            $letter = strtoupper($strLetter[$i]);
            $nOffset = ord($letter) - ord("A") + 1;
            $x += $nOffset * pow($nLength, strlen($strLetter) - 1 - $i);
        }
        return $x;
    }
    private function updateWorkbookLinks()
    {
        unset(
            $this->officeDocument->bookViews[0]->workbookView[0]["activeTab"]
        );

        $this->arrSheets = [];
        $this->arrSheetPath = [];
        //making sheet index
        $ixSheet = 1;
        foreach ($this->officeDocument->sheets->sheet as $sheet) {
            $oldId = (string) $sheet->attributes("r", true)->id;
            $newId = "rId{$ixSheet}";

            $sheet->attributes("r", true)->id = $newId;

            foreach (
                $this->arrXMLs[$this->officeDocumentRelPath]->Relationship
                as $Relationship
            ) {
                if (
                    $oldId == (string) $Relationship["Id"] &&
                    (string) $Relationship["Type"] ==
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
                ) {
                    $Relationship["Id"] = $newId;
                    $oldPath = (string) $Relationship["Target"];
                    if ($oldId != $newId) {
                        $newPath = dirname($oldPath) . "/sheet{$ixSheet}.xml";
                        $Relationship["Target"] = $newPath; //path in relation
                    } else {
                        $newPath = $oldPath;
                    }
                    break;
                }
            }
            if (!$newPath) {
                $newPath = $oldPath = "worksheets/sheet{$ixSheet}.xml";
            }
            $oldAbsolutePath = self::getPathByRelTarget(
                $this->officeDocumentRelPath,
                $oldPath
            );
            $newAbsolutePath = self::getPathByRelTarget(
                $this->officeDocumentRelPath,
                $newPath
            );
            if ($oldId != $newId) {
                $this->renameFile($oldAbsolutePath, $newAbsolutePath);
            }

            $this->arrSheets[(string) $sheet["sheetId"]] =
                &$this->arrXMLs[$newAbsolutePath];
            $this->arrSheetPath[(string) $sheet["sheetId"]] = $newAbsolutePath;

            if ($oldId != $newId) {
                // rename sheet rels only if sheet is changed
                $relPath = self::getRelFilePath($oldAbsolutePath);
                if ($this->arrXMLs[$relPath]->Relationship) {
                    foreach (
                        $this->arrXMLs[$relPath]->Relationship
                        as $Relationship
                    ) {
                        $oldRelTarget = (string) $Relationship["Target"];
                        $newRelTarget = preg_replace(
                            "/([0-9]+)\.([a-z0-9]+)/i",
                            $ixSheet . '.\2',
                            $oldRelTarget
                        );
                        $Relationship["Target"] = $newRelTarget;
                        $this->renameFile(
                            self::getPathByRelTarget($relPath, $oldRelTarget),
                            self::getPathByRelTarget($relPath, $newRelTarget)
                        );
                    };
                }
                $this->renameFile(
                    $relPath,
                    self::getRelFilePath($newAbsolutePath)
                );
            }

            $ixSheet++;
        }
        $ixRel = 0;
        foreach (
            $this->arrXMLs[$this->officeDocumentRelPath]->Relationship
            as $Relationship
        ) {
            if (
                (string) $Relationship["Type"] ==
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
            ) {
                $this->arrXMLs[$this->officeDocumentRelPath]->Relationship[
                    $ixRel
                ]["Id"] = "rId{$ixSheet}";
            }
            if (
                (string) $Relationship["Type"] ==
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
            ) {
                $this->arrXMLs[$this->officeDocumentRelPath]->Relationship[
                    $ixRel
                ]["Id"] = "rId" . ($ixSheet + 1);
            }
            if (
                (string) $Relationship["Type"] ==
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
            ) {
                $this->arrXMLs[$this->officeDocumentRelPath]->Relationship[
                    $ixRel
                ]["Id"] = "rId" . ($ixSheet + 2);
            }
            if (
                (string) $Relationship["Type"] ==
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain"
            ) {
                $this->arrXMLs[$this->officeDocumentRelPath]->Relationship[
                    $ixRel
                ]["Id"] = "rId" . ($ixSheet + 3);
            }
            if (
                (string) $Relationship["Type"] ==
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink"
            ) {
                $this->arrXMLs[$this->officeDocumentRelPath]->Relationship[
                    $ixRel
                ]["Id"] = "rId" . ($ixSheet + 4);
            }
            $ixRel++;
        }

        $this->updateAppXML();
    }
    private function updateAppXML()
    {
        $nSheetsOld = (int) $this->arrXMLs[
            "/docProps/app.xml"
        ]->HeadingPairs->children("vt", true)->vector->variant[1]->i4[0];
        $nAllPartsCount = count(
            $this->arrXMLs["/docProps/app.xml"]->TitlesOfParts->children(
                "vt",
                true
            )->vector[0]
        );
        $nOtherStuffCount = $nAllPartsCount - $nSheetsOld;
        $nSheetsNew = count($this->arrSheets);
        $this->arrXMLs["/docProps/app.xml"]->HeadingPairs->children(
            "vt",
            true
        )->vector->variant[1]->i4[0] = $nSheetsNew;
        for ($i = $nSheetsOld - 1; $i >= 0; $i--) {
            unset(
                $this->arrXMLs["/docProps/app.xml"]->TitlesOfParts->children(
                    "vt",
                    true
                )->vector[0]->lpstr[$i]
            );
        }
        $oParent = $this->arrXMLs["/docProps/app.xml"]->TitlesOfParts->children(
            "vt",
            true
        )->vector[0];
        $domParent = dom_import_simplexml($oParent);
        $insertBefore = @dom_import_simplexml(
            $this->arrXMLs["/docProps/app.xml"]->TitlesOfParts->children(
                "vt",
                true
            )->vector[0]->lpstr[0]
        );
        foreach ($this->officeDocument->sheets->sheet as $sheet) {
            $xmlLpstr = $oParent->addChild(
                "vt:lpstr",
                (string) $sheet["name"],
                "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
            );
            $domInsert = $domParent->ownerDocument->importNode(
                dom_import_simplexml($xmlLpstr),
                true
            );
            if ($insertBefore !== null) {
                $domParent->insertBefore($domInsert, $insertBefore);
            } else {
                $domParent->appendChild($domInsert);
            }
        }

        $attr = $this->arrXMLs["/docProps/app.xml"]->TitlesOfParts
            ->children("vt", true)
            ->vector->attributes("", true);
        $attr["size"] = $nSheetsNew + $nOtherStuffCount;
    }
    private function renameFile($oldName, $newName)
    {
        $this->arrXMLs[$newName] = $this->arrXMLs[$oldName];
        unset($this->arrXMLs[$oldName]);

        foreach (
            $this->arrXMLs["/[Content_Types].xml"]->Override
            as $Override
        ) {
            if ((string) $Override["PartName"] == $oldName) {
                $Override["PartName"] = $newName;
            }
        }
    }
    public function unzipToDirectory($zipFilePath, $targetDirName)
    {
        if (file_exists($targetDirName)) {
            self::rmrf($targetDirName);
        }

        if (!@mkdir($targetDirName, 0777, true)) {
            throw new ITFXlsReader_Exception(
                "Unable to create directory to unpack files"
            );
        }

        if (!file_exists($zipFilePath)) {
            throw new ITFXlsReader_Exception("File not found: {$zipFilePath}");
        }

        $zip = zip_open($zipFilePath);
        if (!$zip) {
            throw new ITFXlsReader_Exception(
                "Wrong file format: {$zipFilePath}"
            );
        }
        try{
            while ($zip_entry = zip_read($zip)) {
                $strFileName =
                    $targetDirName .
                    self::DS .
                    str_replace("/", self::DS, zip_entry_name($zip_entry));
                $dir = dirname($strFileName);
                if (!file_exists($dir)) {
                    mkdir($dir, 0777, true);
                }
                zip_entry_open($zip, $zip_entry);
                $strFile = zip_entry_read(
                    $zip_entry,
                    zip_entry_filesize($zip_entry)
                );
                file_put_contents($strFileName, $strFile);
                unset($strFile);
                zip_entry_close($zip_entry);
            }
            zip_close($zip);
            unset($zip);
        }catch(Exception $e){
            throw new ITFXlsReader_Exception($e);
        }
    }
    private function unzipToMemory($zipFilePath)
    {
        $targetDirName = tempnam(sys_get_temp_dir(), "ITFXlsReader_");

        $this->unzipToDirectory($zipFilePath, $targetDirName);

        $ITFXlsReader_FS = new ITFXlsReader_FS($targetDirName);
        $arrRet = $ITFXlsReader_FS->get();

        self::rmrf($targetDirName);

        return $arrRet;
    }
    protected function rmrf($dir)
    {
        if (is_dir($dir)) {
            $ffs = scandir($dir);
            foreach ($ffs as $file) {
                if ($file == "." || $file == "..") {
                    continue;
                }
                $file = $dir . self::DS . $file;
                if (is_dir($file)) {
                    self::rmrf($file);
                } else {
                    unlink($file);
                }
            }
            rmdir($dir);
        } else {
            unlink($dir);
        }
    }
    public function Output($fileName = "", $dest = "S")
    {
        if ($fileName && func_num_args() === 1) {
            $dest = "D";
        }
        if (
            preg_match("/[" . preg_quote("/\\", "/") . "]/", $fileName) &&
            func_num_args() === 1
        ) {
            $dest = "F";
        }

        if (!$fileName || in_array($dest, ["I", "D"])) {
            $fileNameSrc = $fileName;
            $fileName = tempnam(sys_get_temp_dir(), "ITFXlsReader_");
            $remove = $dest !== "F";
        }

        if (is_writable($fileName) || is_writable(dirname($fileName))) {
            include_once dirname(__FILE__) . ITFXlsReader::DS . "zipfile.php";
            $zip = new zipfile();
            foreach ($this->arrXMLs as $xmlFileName => $fileContents) {
                $zip->addFile(
                    is_object($fileContents)
                        ? $fileContents->asXML()
                        : $fileContents,
                    str_replace("/", self::DS, ltrim($xmlFileName, "/"))
                );
            }
            file_put_contents($fileName, $zip->file());
        } else {
            throw new ITFXlsReader_Exception(
                'could not write to file "' . $fileName . '"'
            );
        }

        switch ($dest) {
            case "I":
            case "D":
                if (ini_get("zlib.output_compression")) {
                    ini_set("zlib.output_compression", "Off");
                }
                header("Pragma: public");
                header("Expires: Sat, 26 Jul 1997 05:00:00 GMT"); // Date in the past
                header("Last-Modified: " . gmdate("D, d M Y H:i:s") . " GMT");
                header("Cache-Control: no-store, no-cache, must-revalidate"); // HTTP/1.1
                header("Cache-Control: pre-check=0, post-check=0, max-age=0"); // HTTP/1.1
                header("Pragma: no-cache");
                header("Expires: 0");
                header("Content-Transfer-Encoding: none");
                header(
                    "Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                );
                if ($dest == "I") {
                    header('Content-Disposition: inline"');
                }
                if ($dest == "D") {
                    $outFileName = $fileNameSrc
                        ? basename($fileNameSrc)
                        : $this->defaultFileName;
                    if (!preg_match('/\.xlsx$/i', $outFileName)) {
                        $outFileName .= ".xlsx";
                    }
                    header(
                        "Content-Disposition: attachment; filename*=UTF-8''" .
                            rawurlencode($outFileName)
                    );
                }
                readfile($fileName);
                unlink($fileName);

                die();
            case "F":
                $r = $fileName;
                break;
            case "S":
            default:
                $r = file_get_contents($fileName);
                break;
        }

        if ($remove) {
            unlink($fileName);
        }

        return $r;
    }
}

class ITFXlsReader_Exception extends Exception
{
    public function __construct($msg)
    {
        parent::__construct("ITFXlsReader error: " . $msg);
    }
    public function __toString()
    {
        return htmlspecialchars($this->getMessage());
    }
}
class ITFXlsReader_FS
{
    private $path;
    public $dirs = [];
    public $filesContent = [];
    public function __construct($path)
    {
        $this->path = rtrim($path, ITFXlsReader::DS);
        return $this;
    }
    public function get()
    {
        $this->_scan(ITFXlsReader::DS);
        return [$this->dirs, $this->filesContent];
    }
    private function _scan($pathRel)
    {
        if ($handle = opendir($this->path . $pathRel)) {
            while (false !== ($item = readdir($handle))) {
                if ($item == ".." || $item == ".") {
                    continue;
                }
                if (is_dir($this->path . $pathRel . $item)) {
                    $this->dirs[] = ltrim($pathRel, ITFXlsReader::DS) . $item;
                    $this->_scan($pathRel . $item . ITFXlsReader::DS);
                } else {
                    $this->filesContent[
                        ltrim($pathRel, ITFXlsReader::DS) . $item
                    ] = file_get_contents($this->path . $pathRel . $item);
                }
            }
            closedir($handle);
        }
    }
}
