import sys
import os
import shutil
from pathlib import Path
import subprocess
import win32com.client

WINMERGE_EXE = "C:\Program Files\WinMerge\WinMergeU.exe"  # WinMergeへのパス


class WinMergeXlsx:
    def __init__(self, folder1, folder2, output_xlsx='./output.xlsx'):
        # 絶対パス取得
        self.folder1 = Path(folder1).absolute()
        self.folder2 = Path(folder2).absolute()
        self.output_xlsx = Path(output_xlsx).absolute()

        # WinMerge出力レポートのパスを準備
        parent = str(self.output_xlsx.parent)
        stem = str(self.output_xlsx.stem)
        self.output_html = Path(parent + '/' + stem + '.html')
        self.output_html_files = Path(parent + '/' + stem + '.files')

    def generate(self):
        # 事前準備
        try:
            if win32com.client.GetObject(Class='Excel.Application'):
                print("\nError : Excelを閉じてください。")
                sys.exit(-1)
        except win32com.client.pywintypes.com_error:
            pass
        if (os.path.exists(self.output_html)):
            try:
                os.remove(self.output_html)
            except PermissionError:
                print("\nError : " + str(self.output_html) + "を閉じてください。")
                sys.exit(-1)
        if (os.path.isdir(self.output_html_files)):
            try:
                shutil.rmtree(self.output_html_files)
            except PermissionError:
                print("\nError : " + str(self.output_html_files) + "を閉じてください。")
                sys.exit(-1)
        if (os.path.exists(self.output_xlsx)):
            try:
                os.remove(self.output_xlsx)
            except PermissionError:
                print("\nError : " + str(self.output_xlsx) + "を閉じてください。")
                sys.exit(-1)

        # 比較結果出力
        self._gen_winmerge_report()
        self._convert_html_to_xlsx()

    def _gen_winmerge_report(self):
        command = [
            WINMERGE_EXE,
            str(self.folder1),                     # 比較元のフォルダ
            str(self.folder2),                     # 比較先のフォルダ
            "/minimize",                           # ウィンドウ最小化で起動
            "/noninteractive",                     # レポート出力後に終了
            "/cfg",
            "Settings/DirViewExpandSubdirs=1",     # 自動的にサブフォルダーを展開する
            "/cfg",
            "ReportFiles/ReportType=2",            # シンプルなHTML形式
            "/cfg",
            "ReportFiles/IncludeFileCmpReport=1",  # ファイル比較レポートを含める
            "/r",                                  # すべてのサブフォルダ内のすべてのファイルを比較
            "/u",                                  # 最近使用した項目リストに追加しない
            "/or",                                 # レポートを出力
            str(self.output_html),                 # レポートのパス
        ]
        print(' '.join(command))
        subprocess.run(command)

    def _convert_html_to_xlsx(self):
        # 差分一覧の読み込み処理
        excel = win32com.client.Dispatch('Excel.Application')
        output_wb = excel.Workbooks.Open(self.output_html)
        output_ws = output_wb.Worksheets(1)
        row = 6
        cell = output_ws.Range('A' + str(row))
        while cell.Value:
            if cell.Hyperlinks.Count > 0:
                sheet_name = cell.Value
                cell2 = output_ws.Range('B' + str(row))
                if cell2.Value is not None:
                    sheet_name = cell2.Value.replace('\\', '_') + '_' + cell.Value
                # 差分ファイルのリネーム
                os.rename(str(self.output_html_files) + '/' + sheet_name + '.html', str(self.output_html_files) + '/' + cell.Value + '.html')
                # ハイパーリンクの編集
                subaddress = cell.Value + '!A1'
                for hl in cell.Hyperlinks:
                    hl.Address = ""
                    hl.SubAddress = subaddress
                print(sheet_name + ' ---> ' + cell.Value)
            row += 1
            cell = output_ws.Range('A' + str(row))

        # ファイル差分のコピー
        count = 1
        for html in self.output_html_files.glob('**/*.html'):
            file_wb = excel.Workbooks.Open(html)
            file_ws = file_wb.Worksheets(1)
            file_ws.Copy(Before = None, After=output_wb.Worksheets(count))
            count += 1

        # 開始セルを移動
        output_ws.Activate()
        output_ws.Range('A1').Select()

        # 結果の保存
        output_wb.SaveAs(str(self.output_xlsx), FileFormat=51)  # xlOpenXMLWorkbook
        excel.Quit()
        print("xlsxへの変換が完了しました。")


if __name__ == '__main__':
    if len(sys.argv) < 3:
        print(f'Usage : {sys.argv[0]} <folder1> <folder2>')
        sys.exit(1)

    WinMergeXlsx(sys.argv[1], sys.argv[2]).generate()
