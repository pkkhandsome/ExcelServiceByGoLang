<?php
    chdir(dirname(__FILE__));
    include("../Phplib/ExcelWriter.php");
    $ExcelWriter = NEW ExcelWriter(dirname(__FILE__)."/test.xlsx");
    if($ExcelWriter->errored)
    {
        echo $ExcelWriter->error_msg;
        Exit;
    }
    echo $ExcelWriter->xlsId."\n";
    
    $ExcelWriter->SetFontSize(9);
    
    $myIndex = $ExcelWriter->AddSheet();
    echo $myIndex."\n";
    $myIndex = $ExcelWriter->GetSheetIndex("NewSheet");
    echo $myIndex."\n";
    
    $myIndex = 0;

    $ExcelWriter->SetSheetName($myIndex,"XSheet");
    
    $ExcelWriter->SetCellValue($myIndex,1,1,"xxxttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttttt");
    $ExcelWriter->SetCellFloatValue($myIndex,1,2,1234567.222);
    $ExcelWriter->SetCellFloatValue($myIndex,1,3,1234567.222);
    $ExcelWriter->SetCellFloatValue($myIndex,2,3,1234567.222);
    $ExcelWriter->SetCellFloatValue($myIndex,5,3,1234567.222);
    $ExcelWriter->SetCellFormula($myIndex,1,4,"=A2+A3");
    
    $ExcelWriter->SetCellImage($myIndex,1,5,"C:\\phpstudy\\eis\\grid48x48.png",48,48);
    
    $ExcelWriter->SetCellHylink($myIndex,1,6,"Sheet1!A1");
    
    //$ExcelWriter->MergeCell($myIndex,1,1,2,1);
    $ExcelWriter->SetRangeBgColor($myIndex,1,1,2,1,ExcelColor::xcRed);
    $ExcelWriter->SetCellWrap($myIndex,1,1,1,1);
    $ExcelWriter->SetRangeBgColor($myIndex,1,1,1,1,ExcelColor::xcWhite);
    
    $ExcelWriter->SetRangeBorderColor($myIndex,5,3,9,10, ExcelBorderType :: cbsThin , 'FF0000');
    $ExcelWriter->SetRangeBorderColor($myIndex,4,2,9,11, ExcelBorderType ::cbsHair , 'FF0000');

    $ExcelWriter->SetRangeBorderColor($myIndex,11,11,11,15, ExcelBorderType ::cbsDashDot , 'FF0000');



    $ExcelWriter->MergeCell($myIndex,7,3,8,3);
    
    $ExcelWriter->SetCellValue($myIndex,7,3,"我们是共产主义接班人");
    $ExcelWriter->SetCellFontStyle($myIndex,7,3,7,3,ExcelFontStyle::xfsBold_Italic_StrikeOut_Underline);
    
    
    $ExcelWriter->SetRangeAlignMent($myIndex,1,1,1,1, ExcelCellHorizAlignment::chaCenter, ExcelCellVertAlignment::cvaCenter);
    $ExcelWriter->SetCellFontSize($myIndex,1,1,1,1,20);
    
    $ExcelWriter->SetCellFontSize($myIndex,1,1,1,1,20);
    
    $ExcelWriter->SetCellFontColor($myIndex,1,1,1,1,"FF0000");
    
    $myStyle = array(
        "bold" => true,
        "italic" => true,
        "size" => 50,
        "color"=>"FF00FF"
    );
    
    $ExcelWriter->SetCellFontJsonStyle($myIndex,1,2,1,2,$myStyle);
    
    $ExcelWriter->AutoCellWidth($myIndex,1,10);
    
    $ExcelWriter->SetColWidth($myIndex,1,30);
    
    $ExcelWriter->SetRowHeight($myIndex,1,15);
    
    $ExcelWriter->SetCellValue($myIndex,15,1,"111");
    $ExcelWriter->SetCellValue($myIndex,5,0,"111");
    
    $myRowCount = $ExcelWriter->GetRowCount($myIndex);
    echo $myRowCount."\n";
    $myColCount = $ExcelWriter->GetColCount($myIndex);
    echo $myColCount."\n";
    
    $ExcelWriter->SaveXls();
?>