# MyScript
Script Collection

## Windows ショートカットを適用する方法
1. ショートカットファイルを以下に格納する
   C:\Users\GN01019\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Shortcut
2. ショートカットのリンク先の先頭に以下を追加する
   powershell -ExecutionPolicy RemoteSigned -File [PowerShellスクリプトのフルパス]
