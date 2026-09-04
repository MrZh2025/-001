$ErrorActionPreference = 'Stop'
$visio = New-Object -ComObject Visio.InvisibleApp
$visio.AlertResponse = 1

try {
    $doc = $visio.Documents.Add("")
    $page = $doc.Pages.Item(1)
    
    # 设置 A4 竖向页面
    $page.PageSheet.CellsU("PageWidth").FormulaU = "8.27 in"
    $page.PageSheet.CellsU("PageHeight").FormulaU = "11.69 in"
    
    # 获取标准流程图模具 (BASFLO_M.VSSX)
    $stencilPath = "C:\Program Files\Microsoft Office\root\Office16\Visio Content\2052\BASFLO_M.VSSX"
    if (-not (Test-Path $stencilPath)) {
        $stencilPath = "C:\Program Files\Microsoft Office\root\Office16\Visio Content\1033\BASFLO_M.VSSX"
    }
    
    $stencil = $visio.Documents.OpenEx($stencilPath, 66)
    
    $processMaster = $stencil.Masters.ItemU("Process")
    $decisionMaster = $stencil.Masters.ItemU("Decision")
    $terminatorMaster = $stencil.Masters.ItemU("Start/End")
    $connectorMaster = $stencil.Masters.ItemU("Dynamic connector")
    
    # 格式化形状函数
    function Format-Shape {
        param($Shape, [string]$Text, [double]$W, [double]$H, [string]$Fill, [string]$Line, [string]$TextColor="RGB(33,37,41)", [double]$FontSize=10.5, [int]$Bold=0)
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

    # 连接两个形状（正交动态连接线）
    function Connect-Shapes {
        param($From, $To, [double]$FromX=0.5, [double]$FromY=0.0, [double]$ToX=0.5, [double]$ToY=1.0, [string]$LineColor="RGB(60,60,60)")
        $conn = $page.Drop($connectorMaster, 0, 0)
        $conn.CellsU("BeginX").GlueToPos($From, $FromX, $FromY)
        $conn.CellsU("EndX").GlueToPos($To, $ToX, $ToY)
        $conn.CellsU("ShapeRouteStyle").FormulaU = "0"
        $conn.CellsU("ConLineRouteExt").FormulaU = "0"
        $conn.CellsU("LineColor").FormulaU = $LineColor
        $conn.CellsU("LineWeight").FormulaU = "1.2 pt"
        $conn.CellsU("EndArrow").FormulaU = "4"
        return $conn
    }
    
    # 1. 大标题
    $title = $page.DrawRectangle(1.0, 10.7, 7.27, 11.3)
    $title.Text = "境内第二类医疗器械注册申报与审批标准流程图"
    $title.CellsU("LinePattern").FormulaU = "0"
    $title.CellsU("FillPattern").FormulaU = "0"
    $title.CellsU("Char.Color").FormulaU = "RGB(15,77,146)"
    $title.CellsU("Char.Size").FormulaU = "14 pt"
    $title.CellsU("Char.Style").FormulaU = "1"
    $title.CellsU("Para.HorzAlign").FormulaU = "1"
    
    # 坐标定义
    $cX = 4.135
    $lX = 1.65
    $rX = 6.62
    
    # 雅致配色定义
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
    
    # 节点1: 向省级药监部门提交注册资料 (Y=9.8)
    $n1 = $page.Drop($processMaster, $cX, 9.8)
    Format-Shape $n1 "向省级药监部门`n提交注册资料" 2.5 0.7 $cBlueFill $cBlueLine "RGB(15,77,146)" 10.5 1
    
    # 节点2: 形式审查 (菱形, Y=8.4)
    $n2 = $page.Drop($decisionMaster, $cX, 8.4)
    Format-Shape $n2 "形式审查`n(5个工作日)" 2.1 1.0 $cGoldFill $cGoldLine "RGB(183,110,0)" 10 1
    
    # 节点3: 形式审查三大分支 (Y=7.0)
    $n3_left = $page.Drop($processMaster, $lX, 7.0)
    Format-Shape $n3_left "发补`n(补正告知)" 1.5 0.6 $cOrangeFill $cOrangeLine "RGB(192,57,43)" 10 0
    
    $n3_mid = $page.Drop($processMaster, $cX, 7.0)
    Format-Shape $n3_mid "受理`n(缴费入卷)" 1.6 0.6 $cBlueFill $cBlueLine "RGB(15,77,146)" 10.5 1
    
    $n3_right = $page.Drop($processMaster, $rX, 7.0)
    Format-Shape $n3_right "不予受理" 1.5 0.6 $cRedFill $cRedLine "RGB(192,57,43)" 10 0
    
    # 节点4: 技术审评 (Y=5.7)
    $n4 = $page.Drop($processMaster, $cX, 5.7)
    Format-Shape $n4 "技术审评`n(省药审中心·60工作日)" 2.4 0.7 $cBlueFill $cBlueLine "RGB(15,77,146)" 10.5 1
    
    # 节点5: 审评三大分支 (Y=4.3)
    $n5_left = $page.Drop($processMaster, $lX, 4.3)
    Format-Shape $n5_left "不予注册`n(终止审批)" 1.5 0.6 $cRedFill $cRedLine "RGB(192,57,43)" 10 0
    
    $n5_mid = $page.Drop($processMaster, $cX, 4.3)
    Format-Shape $n5_mid "注册质量`n体系核查" 2.0 0.7 $cBlueFill $cBlueLine "RGB(15,77,146)" 10.5 1
    
    $n5_right = $page.Drop($processMaster, $rX, 4.3)
    Format-Shape $n5_right "资料补充`n(发补答辩)" 1.5 0.6 $cOrangeFill $cOrangeLine "RGB(192,57,43)" 10 0
    
    # 节点6: 审批决定 (菱形, Y=2.9)
    $n6 = $page.Drop($decisionMaster, $cX, 2.9)
    Format-Shape $n6 "审批决定`n(省局20工作日)" 2.1 1.0 $cGoldFill $cGoldLine "RGB(183,110,0)" 10 1
    
    # 节点7: 发注册证 (Y=1.5)
    $n7 = $page.Drop($terminatorMaster, $cX, 1.5)
    Format-Shape $n7 "发注册证`n(医疗器械注册证+PTR)" 2.4 0.7 $cGreenFill $cGreenLine "RGB(20,90,50)" 11 1
    
    # 连接线绘制
    # 1. 提交资料 -> 形式审查
    Connect-Shapes $n1 $n2 0.5 0.0 0.5 1.0 | Out-Null
    
    # 2. 形式审查 -> 发补 (左折线)
    Connect-Shapes $n2 $n3_left 0.0 0.5 0.5 1.0 | Out-Null
    
    # 3. 形式审查 -> 受理 (中直连)
    Connect-Shapes $n2 $n3_mid 0.5 0.0 0.5 1.0 | Out-Null
    
    # 4. 形式审查 -> 不予受理 (右折线)
    Connect-Shapes $n2 $n3_right 1.0 0.5 0.5 1.0 | Out-Null
    
    # 5. 发补 -> 受理 (横向连线)
    Connect-Shapes $n3_left $n3_mid 1.0 0.5 0.0 0.5 | Out-Null
    
    # 6. 受理 -> 技术审评
    Connect-Shapes $n3_mid $n4 0.5 0.0 0.5 1.0 | Out-Null
    
    # 7. 技术审评 -> 不予注册 (左折线)
    Connect-Shapes $n4 $n5_left 0.0 0.5 0.5 1.0 | Out-Null
    
    # 8. 技术审评 -> 体系核查 (中直连)
    Connect-Shapes $n4 $n5_mid 0.5 0.0 0.5 1.0 | Out-Null
    
    # 9. 技术审评 -> 资料补充 (右折线)
    Connect-Shapes $n4 $n5_right 1.0 0.5 0.5 1.0 | Out-Null
    
    # 10. 资料补充 -> 体系核查 (横向连线)
    Connect-Shapes $n5_right $n5_mid 0.0 0.5 1.0 0.5 | Out-Null
    
    # 11. 体系核查 -> 审批决定
    Connect-Shapes $n5_mid $n6 0.5 0.0 0.5 1.0 | Out-Null
    
    # 12. 审批决定 -> 发注册证
    Connect-Shapes $n6 $n7 0.5 0.0 0.5 1.0 | Out-Null
    
    # 13. 审批决定 -> 不予注册 (不予许可时流转)
    Connect-Shapes $n6 $n5_left 0.0 0.5 0.5 0.0 | Out-Null

    # 输出路径
    $outVsdx = "F:\2026年\梦见2026年\四川省检查员\医疗器械法规汇编知识\二类器械产品注册\第二类医疗器械注册申报审批流程图.vsdx"
    $outPng = "F:\2026年\梦见2026年\四川省检查员\医疗器械法规汇编知识\二类器械产品注册\第二类医疗器械注册申报审批流程图.png"
    
    if (Test-Path $outVsdx) { Remove-Item $outVsdx -Force }
    $doc.SaveAs($outVsdx)
    Write-Host "VSDX_SAVED_SUCCESS"
    
    if (Test-Path $outPng) { Remove-Item $outPng -Force }
    $page.Export($outPng)
    Write-Host "PNG_EXPORTED_SUCCESS"
    
    $stencil.Close()
    $doc.Close()
} catch {
    Write-Host "ERROR: $($_.Exception.Message)"
} finally {
    $visio.Quit()
}
