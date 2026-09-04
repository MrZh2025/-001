# 引入 visio_helpers
$helperPath = "F:\2026年\梦见2026年\开发插件汇总\本地skill\Claude自定义skill\visio-automation\scripts\visio_helpers.ps1"
. $helperPath

$outputPathVsdx = "F:\2026年\梦见2026年\四川省检查员\医疗器械法规汇编知识\二类器械产品注册\第二类医疗器械注册申报审批流程图.vsdx"
$outputPathPng = "F:\2026年\梦见2026年\四川省检查员\医疗器械法规汇编知识\二类器械产品注册\第二类医疗器械注册申报审批流程图.png"

$visio = New-InvisibleVisioApplication
try {
    $doc = $visio.Documents.Add("")
    $page = $doc.Pages.Item(1)
    
    $page.PageSheet.CellsU("PageWidth").FormulaU = "8.27 in"
    $page.PageSheet.CellsU("PageHeight").FormulaU = "11.69 in"
    
    $stencilPath = "C:\Program Files\Microsoft Office\root\Office16\Visio Content\2052\BASFLO_M.VSSX"
    if (-not (Test-Path $stencilPath)) {
        $stencilPath = "C:\Program Files\Microsoft Office\root\Office16\Visio Content\1033\BASFLO_M.VSSX"
    }
    $stencil = Open-VisioStencilReadOnly -Visio $visio -StencilNameOrPath $stencilPath
    
    $processMaster = $stencil.Masters.ItemU("Process")
    $decisionMaster = $stencil.Masters.ItemU("Decision")
    $terminatorMaster = $stencil.Masters.ItemU("Start/End")
    $connectorMaster = $stencil.Masters.ItemU("Dynamic connector")
    
    function Format-Shape {
        param($Shape, [string]$Text, $W, $H, $Fill, $Line, $TextColor="RGB(33,37,41)", $FontSize=10.5, $Bold=0)
        $Shape.Text = $Text
        $Shape.CellsU("Width").FormulaU = "$W in"
        $Shape.CellsU("Height").FormulaU = "$H in"
        $Shape.CellsU("FillForegnd").FormulaU = $Fill
        $Shape.CellsU("LineColor").FormulaU = $Line
        $Shape.CellsU("LineWeight").FormulaU = "1.5 pt"
        $Shape.CellsU("Char.Color").FormulaU = $TextColor
        $Shape.CellsU("Char.Size").FormulaU = "$FontSize pt"
        $Shape.CellsU("Char.Style").FormulaU = "$Bold"
        $Shape.CellsU("Char.Font").FormulaU = "105"
        $Shape.CellsU("Para.HorzAlign").FormulaU = "1"
        $Shape.CellsU("VerticalAlign").FormulaU = "1"
    }
    
    $title = $page.DrawRectangle(1.0, 10.6, 7.27, 11.2)
    $title.Text = "境内第二类医疗器械注册申报与审批标准流程图"
    $title.CellsU("LinePattern").FormulaU = "0"
    $title.CellsU("FillPattern").FormulaU = "0"
    $title.CellsU("Char.Color").FormulaU = "RGB(15,77,146)"
    $title.CellsU("Char.Size").FormulaU = "14 pt"
    $title.CellsU("Char.Style").FormulaU = "1"
    $title.CellsU("Para.HorzAlign").FormulaU = "1"
    
    $cX = 4.135
    $lX = 1.65
    $rX = 6.62
    
    $cBlueFill = "RGB(235, 245, 255)"
    $cBlueLine = "RGB(41, 128, 185)"
    $cGoldFill = "RGB(254, 249, 231)"
    $cGoldLine = "RGB(212, 172, 13)"
    $cGreenFill = "RGB(232, 248, 245)"
    $cGreenLine = "RGB(39, 174, 96)"
    $cOrangeFill = "RGB(255, 243, 224)"
    $cOrangeLine = "RGB(230, 126, 34)"
    $cRedFill = "RGB(253, 237, 236)"
    $cRedLine = "RGB(231, 76, 60)"
    
    $n1 = $page.Drop($processMaster, $cX, 9.7)
    Format-Shape $n1 "向省级药监部门`n提交注册资料" 2.5 0.7 $cBlueFill $cBlueLine "RGB(15,77,146)" 10.5 1
    
    $n2 = $page.Drop($decisionMaster, $cX, 8.3)
    Format-Shape $n2 "形式审查`n(5个工作日)" 2.1 1.0 $cGoldFill $cGoldLine "RGB(183,110,0)" 10 1
    
    $n3_left = $page.Drop($processMaster, $lX, 6.9)
    Format-Shape $n3_left "发补`n(补正告知)" 1.5 0.6 $cOrangeFill $cOrangeLine "RGB(192,57,43)" 10 0
    
    $n3_mid = $page.Drop($processMaster, $cX, 6.9)
    Format-Shape $n3_mid "受理`n(缴费入卷)" 1.6 0.6 $cBlueFill $cBlueLine "RGB(15,77,146)" 10.5 1
    
    $n3_right = $page.Drop($processMaster, $rX, 6.9)
    Format-Shape $n3_right "不予受理" 1.5 0.6 $cRedFill $cRedLine "RGB(192,57,43)" 10 0
    
    $n4 = $page.Drop($processMaster, $cX, 5.6)
    Format-Shape $n4 "技术审评`n(省药审中心·60工作日)" 2.4 0.7 $cBlueFill $cBlueLine "RGB(15,77,146)" 10.5 1
    
    $n5_left = $page.Drop($processMaster, $lX, 4.2)
    Format-Shape $n5_left "不予注册`n(终止审批)" 1.5 0.6 $cRedFill $cRedLine "RGB(192,57,43)" 10 0
    
    $n5_mid = $page.Drop($processMaster, $cX, 4.2)
    Format-Shape $n5_mid "注册质量`n体系核查" 2.0 0.7 $cBlueFill $cBlueLine "RGB(15,77,146)" 10.5 1
    
    $n5_right = $page.Drop($processMaster, $rX, 4.2)
    Format-Shape $n5_right "资料补充`n(发补答辩)" 1.5 0.6 $cOrangeFill $cOrangeLine "RGB(192,57,43)" 10 0
    
    $n6 = $page.Drop($decisionMaster, $cX, 2.8)
    Format-Shape $n6 "审批决定`n(省局20工作日)" 2.1 1.0 $cGoldFill $cGoldLine "RGB(183,110,0)" 10 1
    
    $n7 = $page.Drop($terminatorMaster, $cX, 1.4)
    Format-Shape $n7 "发注册证`n(医疗器械注册证+PTR)" 2.4 0.7 $cGreenFill $cGreenLine "RGB(20,90,50)" 11 1
    
    Connect-VisioShapesOrthogonal -Page $page -ConnectorMaster $connectorMaster -From $n1 -To $n2 -FromX 0.5 -FromY 0.0 -ToX 0.5 -ToY 1.0 | Out-Null
    Connect-VisioShapesOrthogonal -Page $page -ConnectorMaster $connectorMaster -From $n2 -To $n3_left -FromX 0.0 -FromY 0.5 -ToX 0.5 -ToY 1.0 | Out-Null
    Connect-VisioShapesOrthogonal -Page $page -ConnectorMaster $connectorMaster -From $n2 -To $n3_mid -FromX 0.5 -FromY 0.0 -ToX 0.5 -ToY 1.0 | Out-Null
    Connect-VisioShapesOrthogonal -Page $page -ConnectorMaster $connectorMaster -From $n2 -To $n3_right -FromX 1.0 -FromY 0.5 -ToX 0.5 -ToY 1.0 | Out-Null
    Connect-VisioShapesOrthogonal -Page $page -ConnectorMaster $connectorMaster -From $n3_left -To $n3_mid -FromX 1.0 -FromY 0.5 -ToX 0.0 -FromY 0.5 | Out-Null
    Connect-VisioShapesOrthogonal -Page $page -ConnectorMaster $connectorMaster -From $n3_mid -To $n4 -FromX 0.5 -FromY 0.0 -ToX 0.5 -ToY 1.0 | Out-Null
    Connect-VisioShapesOrthogonal -Page $page -ConnectorMaster $connectorMaster -From $n4 -To $n5_left -FromX 0.0 -FromY 0.5 -ToX 0.5 -ToY 1.0 | Out-Null
    Connect-VisioShapesOrthogonal -Page $page -ConnectorMaster $connectorMaster -From $n4 -To $n5_mid -FromX 0.5 -FromY 0.0 -ToX 0.5 -ToY 1.0 | Out-Null
    Connect-VisioShapesOrthogonal -Page $page -ConnectorMaster $connectorMaster -From $n4 -To $n5_right -FromX 1.0 -FromY 0.5 -ToX 0.5 -ToY 1.0 | Out-Null
    Connect-VisioShapesOrthogonal -Page $page -ConnectorMaster $connectorMaster -From $n5_right -To $n5_mid -FromX 0.0 -FromY 0.5 -ToX 1.0 -ToY 0.5 | Out-Null
    Connect-VisioShapesOrthogonal -Page $page -ConnectorMaster $connectorMaster -From $n5_mid -To $n6 -FromX 0.5 -FromY 0.0 -ToX 0.5 -ToY 1.0 | Out-Null
    Connect-VisioShapesOrthogonal -Page $page -ConnectorMaster $connectorMaster -From $n6 -To $n7 -FromX 0.5 -FromY 0.0 -ToX 0.5 -ToY 1.0 | Out-Null
    Connect-VisioShapesOrthogonal -Page $page -ConnectorMaster $connectorMaster -From $n6 -To $n5_left -FromX 0.0 -FromY 0.5 -ToX 0.5 -ToY 0.0 | Out-Null

    if (Test-Path $outputPathVsdx) { Remove-Item $outputPathVsdx -Force }
    $doc.SaveAs($outputPathVsdx)
    Write-Host "VSDX_SAVED_SUCCESS"
    
    if (Test-Path $outputPathPng) { Remove-Item $outputPathPng -Force }
    $page.Export($outputPathPng)
    Write-Host "PNG_EXPORTED_SUCCESS"
    
    $stencil.Close()
    $doc.Close()
} finally {
    $visio.Quit()
}
