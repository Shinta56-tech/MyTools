# アセンブリ========================================================================================

# Add-Type------------------------------------------------------------------------------------
        
# System.Windows.Froms
Add-Type -AssemblyName System.Windows.Forms

# Win32
Add-Type @"
    using System;
    using System.Runtime.InteropServices;
    public class Win32 {
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
    }
    public struct RECT {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }
"@

# Load----------------------------------------------------------------------------------------

# Microsoft.VisualBasic
[void][System.Reflection.Assembly]::Load("Microsoft.VisualBasic, Version=8.0.0.0, Culture=Neutral, PublicKeyToken=b03f5f7f11d50a3a")

# 定数==============================================================================================
    
    # 見出しの表示形式設定
    $HEADLINE_NUMBERFORMATLOCAL = ";;"

# 成果物============================================================================================
    
    # Excelアプリケーションオブジェクト
    $Excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')

    # 見出しセルのXMLオブジェクト
    $HeadLineCellsXML = [xml]"<TreeViewNodes version=""2"" revision=""1""></TreeViewNodes>"

    # ブック目次GUIオブジェクト
    $BookIndexForm = New-Object System.Windows.Forms.Form

# 関数==============================================================================================

    # ブック目次GUIのノードの追加
    function BookIndexAddNode (
            [System.Windows.Forms.TreeView]$treeView
            ,[System.Xml.XmlElement]$element
            ,[System.Windows.Forms.TreeNode]$node
    ){
        # 親ノードが渡されてない
        if ($node -eq $null) {
            $thisNode = $treeView.Nodes.Add($element.GetAttribute("name"),$element.Getattribute("label"))
            foreach ($childElement in $element.Node) {
                BookIndexAddNode -element $childElement -node $thisNode
            }
        # 親ノードが渡されている
        } else {
            $thisNode = $node.Nodes.Add($element.GetAttribute("name"),$element.Getattribute("label"))
            foreach ($childElement in $element.Node) {
                BookIndexAddNode -element $childElement -node $thisNode
            }
        }

        # アクティブシートの見出しノードの場合、展開する
        If ($element.GetAttribute("active") -eq "active") {
            $thisNode.ExpandAll()
            # アクティブシートノードの場合、選択する
            If ($element.GetAttribute("type") -eq "sheet") {
                $treeView.SelectedNode = $thisNode
            }
        }
    }

    #ツリービューのダブルクリック時の処理
    function BookIndex_NodeMouseDoubleClick (
        [System.Windows.Forms.TreeView]$sender
        ,[System.Windows.Forms.TreeNodeMouseClickEventArgs]$e
    ){
        # ノードの展開を維持
        $sender.SelectedNode.ExpandAll()
        # アドレス情報の取得
        $address = $sender.SelectedNode.Name
        $bookName = [RegEx]::Matches($address, "(?<=(\\))[^\\]+?(?=(\]))") | Select-Object -Last 1 | % {$_.Value}
        $sheetName = [RegEx]::Matches($address, "(?<=(')).+?(?=('!))") | Select-Object -Last 1 | % {$_.Value}
        $range = [RegEx]::Matches($address, "(?<=('!)).+") | Select-Object -Last 1 | % {$_.Value}
        # Excelでアドレスを選択して表示
        $book = $Excel.Workbooks($bookName)
        $sheet = $book.Worksheets($sheetName)
        $sheet.Activate()
        $sheet.Range($range).select()
        # Excelをアクティブ
        [Microsoft.VisualBasic.Interaction]::AppActivate($bookName)
    }

# メイン処理========================================================================================

    # ActiveBookの目次XMLの作成---------------------------------------------------------------
        
        # Excelの検索設定
        $Excel.FindFormat.Clear()
        $Excel.FindFormat.NumberFormatLocal = $HEADLINE_NUMBERFORMATLOCAL

        # Excelの全シートループ
        foreach ($sheet in $Excel.ActiveWorkbook.WorkSheets) {

            # XMLシート要素の作成・追加
            $sheetElement = $HeadLineCellsXML.CreateElement("Node")
            $sheetElement.SetAttribute("type", "sheet")
            $sheetElement.SetAttribute("name","[" + $Excel.ActiveWorkbook.FullName + "]'"  + $sheet.Name + "'!A1")
            $sheetElement.SetAttribute("label", $sheet.Name)
            $HeadLineCellsXML.TreeViewNodes.AppendChild($sheetElement)
            
            # アクティブシートの要素フラッグ
            $activeFlag = $false
            If ($Excel.ActiveSheet.Name -eq $sheet.Name) {
                $sheetElement.SetAttribute("active", "active")
                $activeFlag = $true
            }

            # 最初の書式検索
            $foundCell = $sheet.Cells.Find(
                "?*" # What
                ,$sheet.Cells(1,1) # After
                ,-4163 # LookIn
                ,1 # LookAt
                ,1 # SeachOrder
                ,1 # SearchDirection
                ,$true # MatchCase
                ,$true # MatchByte
                ,$true # SearchFormat
                )

            # 該当なしの場合、次のシートへ
            If ($foundCell -eq $null) {
                continue
            }

            # 初回アドレスの格納
            $firstFoundAddress = $foundCell.Address()

            # XMLに格納
            Do {

                # XML見出し要素の作成
                $headLineElement = $HeadLineCellsXML.CreateElement("Node")
                $headLineElement.SetAttribute("name","[" + $Excel.ActiveWorkbook.FullName + "]'" + $sheet.Name + "'!" + $foundCell.Address())
                $headLineElement.SetAttribute("label", $foundCell.Value())
                If ($activeFlag) {$headLineElement.SetAttribute("active", "active")}

                # XML見出し要素の追加
                Switch($foundCell.Style.Name())
                {
                    "見出し 1" {
                        $headLineElement.SetAttribute("type", "headLine1")
                        $sheetElement.AppendChild($headLineElement)
                    }
                    "見出し 2" {
                        $headLineElement.SetAttribute("type", "headLine2")
                        $sheetElement.SelectSingleNode("//Node[@type='headLine1'][last()]").AppendChild($headLineElement)
                    }
                    "見出し 3" {
                        $headLineElement.SetAttribute("type", "headLine3")
                        $sheetElement.SelectSingleNode("//Node[@type='headLine2'][last()]").AppendChild($headLineElement)
                    }
                    "見出し 4" {
                        $headLineElement.SetAttribute("type", "headLine4")
                        $sheetElement.SelectSingleNode("//Node[@type='headLine3'][last()]").AppendChild($headLineElement)
                    }
                }

                # 次の書式検索
                $foundCell = $sheet.Cells.Find(
                    "?*" # What
                    ,$foundCell # After
                    ,-4163 # LookIn
                    ,1 # LookAt
                    ,1 # SeachOrder
                    ,1 # SearchDirection
                    ,$true # MatchCase
                    ,$true # MatchByte
                    ,$true # SearchFormat
                    )

            # 初回アドレスと次見出しセルのアドレスが異なる場合に続行
            } While($firstFoundAddress -ne $foundCell.Address())
        }

        # Excelの検索設定のクリア
        $Excel.FindFormat.Clear()

    # FromBookIndexの作成----------------------------------------------------------------------
        
        # アクティブウィンドウの情報を取得
        $rect = New-Object RECT
        [Win32]::GetWindowRect($Excel.ActiveWindow.HWND, [ref]$rect) | Out-Null

        # GUIの設定
        $BookIndexForm.Text = "ブック目次 - " + $Excel.ActiveWorkbook.Name
        $BookIndexForm.MaximizeBox = $false  # 最大化ボタンを無効にする
        $BookIndexForm.TopMost = $true # 最前面に表示
        $BookIndexForm.Width = 300
        $BookIndexForm.Height = $rect.Bottom - $rect.Top - 100
        $BookIndexForm.StartPosition = "Manual"
        $BookIndexForm.Left = $rect.Right - $BookIndexForm.Width - 13
        $BookIndexForm.Top = $rect.Bottom - $BookIndexForm.Height - 35
        $BookIndexForm.FormBorderStyle = 'Fixed3D'
        #TreeViewの設定
        $treeView = New-Object System.Windows.Forms.TreeView
        $treeView.Location = "10,10"
        $treeView.Width = $BookIndexForm.Width - 40
        $treeView.Height = $BookIndexForm.Height - 70
        [void]$BookIndexForm.Controls.AddRange($treeView)

        # ノードの追加
        foreach ($element in $HeadLineCellsXML.TreeViewNodes.Node) {
            BookIndexAddNode -treeView $treeView -element $element
        }

        # ノードの文字大きさの調整
        $fontSize = 12
        $treeView.Font = New-Object System.Drawing.Font("Arial", $fontSize)
        $treeView.Nodes | ForEach-Object { $_.NodeFont = New-Object System.Drawing.Font("Arial", $fontSize) }

        #ツリービューのダブルクリックイベントの設定
        $treeView.Add_NodeMouseDoubleClick(${function:BookIndex_NodeMouseDoubleClick})

    # FromBookIndexの表示----------------------------------------------------------------------

        # GUIの表示
        [void]$BookIndexForm.ShowDialog()

# 終了処理==========================================================================================
    
    # Excelオブジェクトを消去
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)
