<?php

class ExcelColor
{
    const xc0 = "000000";
    const xc1 = "000000";
    const xc2 = "000000";
    const xc3 = "000000";
    const xc4 = "000000";
    const xc5 = "000000";
    const xc6 = "000000";
    const xc7 = "000000";
    const xcBlack = "000000";
    const xcWhite = "FFFFFF";
    const xcRed = "FF0000";
    const xcBrightGreen = "00FF00";
    const xcBlue = "0000FF";
    const xcYellow = "FFFF00";
    const xcPink = "FF00FF";
    const xcTurquoise = "00FFFF";
    const xcDarkRed = "800000";
    const xcGreen = "008000";
    const xcDarkBlue = "000080";
    const xcBrownGreen = "808000";
    const xcViolet = "800080";
    const xcBlueGreen = "008080";
    const xcGray25 = "C0C0C0";
    const xcGray50 = "808080";
    const xc24 = "9999FF";
    const xc25 = "993366";
    const xc26 = "FFFFCC";
    const xc27 = "CCFFFF";
    const xc28 = "660066";
    const xc29 = "FF8080";
    const xc30 = "0066CC";
    const xc31 = "CCCCFF";
    const xc32 = "000080";
    const xc33 = "FF00FF";
    const xc34 = "FFFF00";
    const xc35 = "00FFFF";
    const xc36 = "800080";
    const xc37 = "800000";
    const xc38 = "008080";
    const xc39 = "0000FF";
    const xcSky = "00CCFF";
    const xcPaleTurquois = "CCFFFF";
    const xcPaleGreen = "CCFFCC";
    const xcLightYellow = "FFFF99";
    const xcPaleSky = "99CCFF";
    const xcRose = "FF99CC";
    const xcLilac = "CC99FF";
    const xcLightBrown = "FFCC99";
    const xcDarkSky = "3366FF";
    const xcDarkTurquois = "33CCCC";
    const xcGrass = "99CC00";
    const xcGold = "FFCC00";
    const xcLightOrange = "FF9900";
    const xcOrange = "FF6600";
    const xcDarkBlueGray = "666699";
    const xcGray40 = "969696";
    const xcDarkGreenGray = "003366";
    const xcEmerald = "339966";
    const xcDarkGreen = "003300";
    const xcOlive = "333300";
    const xcBrown = "993300";
    const xcCherry = "993366";
    const xcIndigo = "333399";
    const xcGray80 = "333333";
    const xcAutomatic = "00000F";
}

class ExcelBorderType
{
    const cbsNone = 0;
    const cbsThin = 1;
    const cbsMedium = 2;
    const cbsDashed = 3;
    const cbsDotted = 4;
    const cbsThick = 5;
    const cbsDouble = 6;
    const cbsHair = 7;
    const cbsMediumDashed = 8;
    const cbsDashDot = 9;
    const cbsMediumDashDot = 10;
    const cbsDashDotDot = 11;
    const cbsMediumDashDotDot = 12;
    const cbsSlantedDashDot = 13;
}

class ExcelCellHorizAlignment
{
    const chaGeneral = "";
    const chaLeft = "left";
    const chaCenter = "center";
    const chaRight = "right";
    const chaFill = "fill";
    const chaJustify = "justify";
    const chaCenterContinuous = "centerContinuous";
    const chaDistributed = "distributed";
}

class ExcelCellVertAlignment
{
    const cvaTop = "top";
    const cvaCenter = "center";
    const cvaBottom = "bottom";
    const cvaJustify = "justify";
    const cvaDistributed = "distributed";
}


class ExcelFontStyle
{
    const xfsNone = 0;
    const xfsBold = 1;
    const xfsItalic = 2;
    const xfsStrikeOut = 3;
    const xfsBold_Italic = 4;
    const xfsBold_StrikeOut = 5;
    const xfsItalic_StrikeOut = 6;
    const xfsBold_Italic_StrikeOut = 7;
    const xfsUnderline = 10;
    const xfsBold_Underline = 11;
    const xfsItalic_Underline = 12;
    const xfsStrikeOut_Underline = 13;
    const xfsBold_Italic_Underline = 14;
    const xfsBold_StrikeOut_Underline = 15;
    const xfsItalic_StrikeOut_Underline = 16;
    const xfsBold_Italic_StrikeOut_Underline = 17;
}

class ExcelWriter
{
    private $skt;
    public $xlsId;
    public $error_msg;
    public $errored = false;
    public $style = array(
        "font"=>array(
            "size"=>9,
            "family"=>"Microsoft Yahei"
        )
    );
    public $fontsizeRati = 8.3;
    public $ponsizeRati = 1.3;

    function __construct($path = '', $version = "")
    {
        if ($path)
        {
            $this->CreateXls($path, $version);
            $this->SetCellDefaultStyle();
        }
    }

    function CreateXls($path, $version)
    {
        //默认只有一个Sheet
        if (file_exists($path))
        {
            if (!unlink($path))
            {
                $this->errored = true;
                $this->error_msg = "file is locked";
                return false;
            }
        }
        $myWindowsPath = iconv("utf-8","gbk",$path);
        copy(dirname(__file__) . "/demo$version.xlsx", $myWindowsPath);
        $this->skt = socket_create(AF_INET, SOCK_STREAM, SOL_TCP);
        if (!$this->skt)
        {
            $this->errored = true;
            $this->error_msg = "socket create error";
            return false;
        }
        $result = @socket_connect($this->skt, '127.0.0.1', 17834);
        if ($result)
        {
            @socket_recv($this->skt, $buf, 100, 0);
            $this->xlsId = trim($buf);
            $mySendStr = $this->xlsId . "#LoadXls#" . $path . "\0";
            @socket_write($this->skt, $mySendStr, strlen($mySendStr));
        } else
        {
            $this->errored = true;
            $this->error_msg = "socket connect error";
            @socket_close($socket);
            return false;
        }
        return $this;
    }

    function AddSheet($SheetName = 'NewSheet')
    {
        $mySendStr = $this->xlsId . "#AddSheet#" . base64_encode($SheetName) . "\0";
        @socket_write($this->skt, $mySendStr, strlen($mySendStr));
        @socket_recv($this->skt, $buf, 100, 0);
        $myIndex = $buf;
        return trim($myIndex);
    }

    function GetSheetIndex($SheetName)
    {
        $mySendStr = $this->xlsId . "#GetSheetIndex#" . base64_encode($SheetName) . "\0";
        @socket_write($this->skt, $mySendStr, strlen($mySendStr));
        @socket_recv($this->skt, $buf, 100, 0);
        $myValue = $buf;
        return trim($myValue);
    }

    function SetSheetName($SheetIndex, $SheetName)
    {
        // $SheetIndex 从0开始
        $mySendStr = $this->xlsId . "#SetSheetName#" . $SheetIndex . "#" . base64_encode($SheetName) .
            "\0";
        @socket_write($this->skt, $mySendStr, strlen($mySendStr));
    }
    
    function SetRowOutlineLevel($SheetIndex, $rowIndex, $Level)
    {
        // $SheetIndex 从0开始
        $mySendStr = $this->xlsId . "#SetRowOutlineLevel#" . $SheetIndex . "#" . $rowIndex . "#" .
            $Level . "\0";
        @socket_write($this->skt, $mySendStr, strlen($mySendStr));
    }

    function SetCellValue($SheetIndex, $Col, $Row, $Value, $ForceString = false)
    {
        if ($Value !== null)
        {
            if(!$ForceString)
            {
                if (is_numeric($Value))
                {
                    $this->SetCellFloatValue($SheetIndex, $Col, $Row, $Value + 0);
                    return true;
                }
            }
            if ($Value == "")
            {
                return false;
            }
            $Value = str_replace("&nbsp;", " ", $Value);
            $mySendStr = $this->xlsId . "#SetCellValue#" . $SheetIndex . "#" . $Col . "#" .
                $Row . "#" . base64_encode($Value) . "\0";
            @socket_write($this->skt, $mySendStr, strlen($mySendStr));
            return true;
        }
        return false;
    }

    function SetCellFloatValue($SheetIndex, $Col, $Row, $Value)
    {
        if (!is_numeric($Value))
        {
            return false;
        }
        if ($Value !== null)
        {
            $mySendStr = $this->xlsId . "#SetCellFloatValue#" . $SheetIndex . "#" . $Col .
                "#" . $Row . "#" . ($Value + 0) . "\0";
            @socket_write($this->skt, $mySendStr, strlen($mySendStr));
            return true;
        }
        return false;
    }

    function SetCellFormula($SheetIndex, $Col, $Row, $Value)
    {
        if ($Value !== null)
        {
            $mySendStr = $this->xlsId . "#SetCellFormula#" . $SheetIndex . "#" . $Col . "#" .
                $Row . "#" . base64_encode($Value) . "\0";
            @socket_write($this->skt, $mySendStr, strlen($mySendStr));
            return true;
        }
        return false;
    }

    function SetCellImage($SheetIndex, $Col, $Row, $Value, $Width, $Height)
    {
        if ($Value !== null)
        {
            // 文字宽度
            $CellWidth = $Width / $this->fontsizeRati;
            // 磅
            $CellHeight = $Height / $this->ponsizeRati;

            $mySendStr = $this->xlsId . "#SetCellImage#" . $SheetIndex . "#" . $Col . "#" .
                $Row . "#" . base64_encode($Value) . "#" . $CellWidth . "#" . $CellHeight ."#" . $Width . "#" . $Height . "\0";
            @socket_write($this->skt, $mySendStr, strlen($mySendStr));
            return true;
        }
        return false;
    }

    function SetCellHylink($SheetIndex, $Col, $Row, $Value, $Type = "Location"
        /*Location or External*/ )
    {
        $mySendStr = $this->xlsId . "#SetCellHylink#" . $SheetIndex . "#" . $Col . "#" .
            $Row . "#" . base64_encode($Value) . "#{$Type}\0";
        @socket_write($this->skt, $mySendStr, strlen($mySendStr));
    }

    function MergeCell($SheetIndex, $StartCol, $StartRow, $EndCol, $EndRow)
    {
        $mySendStr = $this->xlsId . "#MergeCell#" . $SheetIndex . "#" . $StartCol . "#" .
            $StartRow . "#" . $EndCol . "#" . $EndRow . "\0";
        @socket_write($this->skt, $mySendStr, strlen($mySendStr));

    }

    function SetCellWrap($SheetIndex, $StartCol, $StartRow, $EndCol, $EndRow)
    {
        $mySendStr = $this->xlsId . "#SetCellWrap#" . $SheetIndex . "#" . $StartCol .
            "#" . $StartRow . "#" . $EndCol . "#" . $EndRow . "\0";
        @socket_write($this->skt, $mySendStr, strlen($mySendStr));
    }

    function SetRangeBgColor($SheetIndex, $StartCol, $StartRow, $EndCol, $EndRow, $Color =
        ExcelColor::xcWhite)
    {
        //范围内的背景颜色
        $mySendStr = $this->xlsId . "#SetRangeBgColor#" . $SheetIndex . "#" . $StartCol .
            "#" . $StartRow . "#" . $EndCol . "#" . $EndRow . "#" . $Color . "\0";
        @socket_write($this->skt, $mySendStr, strlen($mySendStr));
    }

    function SetRangeBorderColor($SheetIndex, $StartCol, $StartRow, $EndCol, $EndRow,
        $BorderType = ExcelBorderType::cbsThin, $Color = ExcelColor::xcBlack)
    {
        //只画范围外框

        $mySendStr = $this->xlsId . "#SetRangeBorderColor#" . $SheetIndex . "#" . $StartCol .
            "#" . $StartRow . "#" . $EndCol . "#" . $EndRow . "#" . $BorderType . "#" . $Color .
            "\0";
        @socket_write($this->skt, $mySendStr, strlen($mySendStr));

    }

    function SetRangeAllBorderColor($SheetIndex, $StartCol, $StartRow, $EndCol, $EndRow,
        $BorderType = ExcelBorderType::cbsThin, $Color = ExcelColor::xcBlack)
    {
        //范围内的全部画外框和内框

        $mySendStr = $this->xlsId . "#SetRangeAllBorderColor#" . $SheetIndex . "#" . $StartCol .
            "#" . $StartRow . "#" . $EndCol . "#" . $EndRow . "#" . $BorderType . "#" . $Color .
            "\0";
        @socket_write($this->skt, $mySendStr, strlen($mySendStr));

    }

    function SetCellAlignMent($SheetIndex, $Col, $Row, $HorizAlignment =
        ExcelCellHorizAlignment::chaLeft, $VertAlignment = ExcelCellVertAlignment::
        cvaCenter)
    {
        //设置单元格的对其方式
        $mySendStr = $this->xlsId . "#SetCellAlignMent#" . $SheetIndex . "#" . $Col .
            "#" . $Row . "#" . $Col . "#" . $Row . "#" . $HorizAlignment . "#" . $VertAlignment .
            "\0";
        @socket_write($this->skt, $mySendStr, strlen($mySendStr));
    }

    function SetRangeAlignMent($SheetIndex, $StartCol, $StartRow, $EndCol, $EndRow,
        $HorizAlignment = ExcelCellHorizAlignment::chaLeft, $VertAlignment =
        ExcelCellVertAlignment::cvaCenter)
    {
        //设置单元格的对其方式
        $mySendStr = $this->xlsId . "#SetCellAlignMent#" . $SheetIndex . "#" . $StartCol .
            "#" . $StartRow . "#" . $EndCol . "#" . $EndRow . "#" . $HorizAlignment . "#" .
            $VertAlignment . "\0";
        @socket_write($this->skt, $mySendStr, strlen($mySendStr));
    }

    function SetFontSize($FontSize) //整体 Fontsize pt
    {
        $myStyle = $this->style;
        $myStyle["font"]["size"] = $FontSize;
        $this->SetCellDefaultStyle($myStyle);
    }
    
    private function doMergeStyle(&$OldStyle,&$NewStyle)
    {
        foreach($NewStyle as $k=>$v)
        {
            if(key_exists($k,$OldStyle))
            {
                if(is_object($v))
                {
                    $this->doMergeStyle($OldStyle[$k],$v);
                }
                else
                {
                    $OldStyle[$k] = $v;  
                }
            }
            else
            {
                $OldStyle[$k] = $v;
            }
        }
    }
    
    function SetCellDefaultStyle($Style=array())
    {
        foreach($Style as $k=>$v)
        {
            if(key_exists($k,$this->style))
            {
                if(is_object($v))
                {
                    $this->doMergeStyle($this->style[$k],$v);
                }
                else
                {
                    $this->style[$k] = $v;  
                }
            }
            else
            {
                $this->style[$k] = $v;
            }
        }
        $mySendStr = $this->xlsId . "#SetDefaultStyle#" . base64_encode(json_encode($this->style)) . "\0";
        @socket_write($this->skt, $mySendStr, strlen($mySendStr));
    }

    function SetCellFontSize($SheetIndex, $StartCol, $StartRow, $EndCol, $EndRow, $FontSize)
    {
        // 单元格字体大小
        $mySendStr = $this->xlsId . "#SetCellFontSize#" . $SheetIndex . "#" . $StartCol .
            "#" . $StartRow . "#" . $EndCol . "#" . $EndRow . "#" . $FontSize . "\0";
        @socket_write($this->skt, $mySendStr, strlen($mySendStr));
    }

    function SetCellFontColor($SheetIndex, $StartCol, $StartRow, $EndCol, $EndRow, $Color =
        '000000')
    {
        //单元格字体颜色
        $mySendStr = $this->xlsId . "#SetCellFontColor#" . $SheetIndex . "#" . $StartCol .
            "#" . $StartRow . "#" . $EndCol . "#" . $EndRow . "#" . $Color . "\0";
        @socket_write($this->skt, $mySendStr, strlen($mySendStr));
    }

    function SetCellFontStyle($SheetIndex, $StartCol, $StartRow, $EndCol, $EndRow, $Style =
        ExcelFontStyle::xfsNone)
    {
        // 设置字体样式 粗,斜,删除线
        $mySendStr = $this->xlsId . "#SetCellFontStyle#" . $SheetIndex . "#" . $StartCol .
            "#" . $StartRow . "#" . $EndCol . "#" . $EndRow . "#" . $Style . "\0";
        @socket_write($this->skt, $mySendStr, strlen($mySendStr));
    }

    function SetCellFontJsonStyle($SheetIndex, $StartCol, $StartRow, $EndCol, $EndRow,
        $Style = array())
    {
        // {"bold":true,"italic":true,"underline":"single","family":"","color":"000000","size":12}
        // 设置字体样式 粗,斜,删除线
        $myStyle = base64_encode(json_encode($Style));
        $mySendStr = $this->xlsId . "#SetCellFontJsonStyle#" . $SheetIndex . "#" . $StartCol .
            "#" . $StartRow . "#" . $EndCol . "#" . $EndRow . "#" . $myStyle . "\0";
        @socket_write($this->skt, $mySendStr, strlen($mySendStr));
    }

    function AutoCellWidth($SheetIndex, $StartCol, $EndCol)
    {
        //自动调整列宽
        $mySendStr = $this->xlsId . "#AutoCellWidth#" . $SheetIndex . "#" . $StartCol .
            "#" . $EndCol . "\0";
        @socket_write($this->skt, $mySendStr, strlen($mySendStr));
        //由于调整列宽 可能耗时较长 所以此处需要等待返回值
        @socket_recv($this->skt, $buf, 100, 0);
        $myValue = $buf;
        return trim($myValue);
    }

    function AutoCellHeight($SheetIndex, $StartRow, $EndRow)
    {
        //自动行高 默认是自动行高 此功能是无用的 如需去除自动行高
        //则使用 SetRowHeight
        $mySendStr = $this->xlsId . "#AutoCellHeight#" . $SheetIndex . "#" . $StartRow .
            "#" . $EndRow . "\0";
        @socket_write($this->skt, $mySendStr, strlen($mySendStr));
    }

    function SetRowHeight($SheetIndex, $StartRow, $height = 25)
    {
        //自动行高 默认是自动行高 此功能是无用的 如需去除自动行高
        //则使用 SetRowHeight
        $mySendStr = $this->xlsId . "#SetRowHeight#" . $SheetIndex . "#" . $StartRow .
            "#" . $height . "\0";
        @socket_write($this->skt, $mySendStr, strlen($mySendStr));
    }

    function SetColWidth($SheetIndex, $Col, $width)
    {
        // 单位是字体个数
        $mySendStr = $this->xlsId . "#SetColWidth#" . $SheetIndex . "#" . $Col . "#" . $width .
            "\0";
        @socket_write($this->skt, $mySendStr, strlen($mySendStr));
    }
    
    function SetColsWidth($SheetIndex, $colFrom, $colEnd, $width)
    {
        //"10,20,30,40"
        $widthArr = explode(",",$width);
        for($i=$colFrom;$i<=$colEnd;$i++)
        {
            $myIndex = $i - $colFrom;
            if(count($widthArr)<=$myIndex)
            {
                continue;
            }
            if(!$widthArr[$myIndex])
            {
                continue;
            }
            
            // 单位是字体个数
            $mySendStr = $this->xlsId . "#SetColWidth#" . $SheetIndex . "#" . $i . "#" . $widthArr[$myIndex] .
                "\0";
            @socket_write($this->skt, $mySendStr, strlen($mySendStr));
        }
    }

    function SetColWidthPixel($SheetIndex, $Col, $width)
    {
        $width = $width / $this->fontsizeRati;
        $mySendStr = $this->xlsId . "#SetColWidth#" . $SheetIndex . "#" . $Col . "#" . $width .
            "\0";
        @socket_write($this->skt, $mySendStr, strlen($mySendStr));
    }

    function GetRowCount($SheetIndex)
    {
        $mySendStr = $this->xlsId . "#GetRowCount#" . $SheetIndex . "\0";
        @socket_write($this->skt, $mySendStr, strlen($mySendStr));
        @socket_recv($this->skt, $buf, 100, 0);
        $myValue = $buf;
        return trim($myValue);
    }

    function GetColCount($SheetIndex)
    {
        $mySendStr = $this->xlsId . "#GetColCount#" . $SheetIndex . "\0";
        @socket_write($this->skt, $mySendStr, strlen($mySendStr));
        @socket_recv($this->skt, $buf, 100, 0);
        $myValue = $buf;
        return trim($myValue);
    }

    function SaveXls()
    {
        $mySendStr = $this->xlsId . "#SaveXls\0";
        @socket_write($this->skt, $mySendStr, strlen($mySendStr));
        @socket_recv($this->skt, $buf, 100, 0);
    }

    function __destruct()
    {
        if ($this->xlsId)
        {
            $mySendStr = $this->xlsId . "#close\0";
            @socket_write($this->skt, $mySendStr, strlen($mySendStr));
            @socket_recv($this->skt, $buf, 100, 0);
            @socket_close($this->skt);
        }
    }
}
