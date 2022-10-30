import sys
import os
import shutil
from pathlib import Path
import subprocess
import win32com.client

WINMERGE_EXE = 'C:\Program Files\WinMerge\WinMergeU.exe'  # WinMergeへのパス
WINMERGE_OPTIONS = [
    '/minimize',                           # ウィンドウ最小化で起動
    '/noninteractive',                     # レポート出力後に終了
    '/cfg',
    'Settings/DirViewExpandSubdirs=1',     # 自動的にサブフォルダーを展開する
    '/cfg',
    'ReportFiles/ReportType=2',            # シンプルなHTML形式
    '/cfg',
    'ReportFiles/IncludeFileCmpReport=1',  # ファイル比較レポートを含める
    '/r',                                  # すべてのサブフォルダ内のすべてのファイルを比較
    '/u',                                  # 最近使用した項目リストに追加しない
    '/or',                                 # レポートを出力
]

xlUp = -4162
xlOpenXMLWorkbook = 51

SUMMARY_WS_NUM = 1        # 一覧シートのワークシート番号
SUMMARY_START_ROW = 6     # 一覧シートの表の開始行
SUMMARY_NAME_COL = 'A'    # 一覧シートの表の名前列
SUMMARY_FOLDER_COL = 'B'  # 一覧シートの表のフォルダー列

HOME_POSITION = 'A1' # ホームポジション

DIFF_START_ROW = 2                                  # 差分シートの開始行
DIFF_ZOOM_RATIO = 85                                # 差分シートのズームの倍率
DIFF_FORMATS = {                                    # 差分シートの書式設定
    'no': [                                         # 行番号列
        {'range': 'A1', 'width': 3},                # 左側
        {'range': 'C1', 'width': 3},                # 右側
    ],
    'code': [                                       # ソースコード列
        {'range': 'B:B', 'font': 'ＭＳ ゴシック'},  # 左側
        {'range': 'D:D', 'font': 'ＭＳ ゴシック'},  # 右側
    ],
}


class WinMergeXlsx:
    def __init__(self, base, latest, output='./output.xlsx'):
        self.base = Path(base).absolute()
        self.latest = Path(latest).absolute()
        self.output = Path(output).absolute()

        parent = str(self.output.parent)
        stem = str(self.output.stem)
        self.output_html = Path(parent + '/' + stem + '.html')
        self.output_html_files = Path(parent + '/' + stem + '.files')

    def generate(self):
        self._setup()
        self._generate_html_by_winmerge()
        self._convert_html_to_xlsx()

    def _setup(self):
        self._setup_excel_application()
        self._setup_output_files()

    def _setup_excel_application(self):
        try:
            if win32com.client.GetObject(Class='Excel.Application'):
                self._message_and_exit('Excelを閉じてください。')
        except win32com.client.pywintypes.com_error:
            pass

    def _setup_output_files(self):
        if (os.path.exists(self.output_html)):
            try:
                os.remove(self.output_html)
            except PermissionError:
                self._message_and_exit(str(self.output_html) + 'へのアクセス権がありません。')
        if (os.path.isdir(self.output_html_files)):
            try:
                shutil.rmtree(self.output_html_files)
            except PermissionError:
                self._message_and_exit(str(self.output_html_files) + 'へのアクセス権がありません。')
        if (os.path.exists(self.output)):
            try:
                os.remove(self.output)
            except PermissionError:
                self._message_and_exit(str(self.output) + 'へのアクセス権がありません。')

    def _message_and_exit(self, message):
        print('\nError : ' + message)
        sys.exit(-1)

    def _generate_html_by_winmerge(self):
        command = [
            WINMERGE_EXE,
            str(self.base),         # 比較元のフォルダ
            str(self.latest),       # 比較先のフォルダ
            *WINMERGE_OPTIONS,      # WinMergeのコマンドライン実行オプション
            str(self.output_html),  # レポートのパス
        ]
        print(' '.join(command))
        subprocess.run(command)

    def _convert_html_to_xlsx(self):
        try:
            self._open_book()
            self._format_summary_sheet()
            self._copy_html_files()
            self._format_diff_sheets()
            self._set_home_position()
            self._save_book()

        finally:
            self.excel.Quit()

    def _open_book(self):
        self.excel = win32com.client.Dispatch('Excel.Application')
        self.wb = self.excel.Workbooks.Open(self.output_html)
        self.summary_ws = self.wb.Worksheets(SUMMARY_WS_NUM)

    def _format_summary_sheet(self):
        ws = self.summary_ws
        end_row = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        for row in range(SUMMARY_START_ROW, end_row+1):
            name_cell = ws.Range(SUMMARY_NAME_COL + str(row))
            if not name_cell.Value:
                break
            if name_cell.Hyperlinks.Count > 0:
                self._change_hyperlink(name_cell)
                folder_cell = ws.Range(SUMMARY_FOLDER_COL + str(row))
                if folder_cell.Value:
                    self._rename_html_files(name_cell, folder_cell)

    def _rename_html_files(self, name_cell, folder_cell):
        sheet_name = folder_cell.Value.replace('\\', '_') + '_' + name_cell.Value
        src = f'{self.output_html_files}/{sheet_name}.html'
        dst = f'{self.output_html_files}/{name_cell.Value}.html'
        os.rename(src, dst)
        print(sheet_name + ' ---> ' + name_cell.Value)

    def _change_hyperlink(self, name_cell):
        for hl in name_cell.Hyperlinks:
            hl.Address = ''
            hl.SubAddress = name_cell.Value + '!' + HOME_POSITION

    def _copy_html_files(self):
        for count, html in enumerate(self.output_html_files.glob('**/*.html'), 1):
            diff_wb = self.excel.Workbooks.Open(html)
            diff_ws = diff_wb.Worksheets(1)
            diff_ws.Copy(Before=None, After=self.wb.Worksheets(count))

    def _format_diff_sheets(self):
        for i in range(DIFF_START_ROW, self.wb.Worksheets.Count+1):
            ws = self.wb.Worksheets(i)
            self._set_zoom(ws)
            self._format_no_cols(ws)
            self._format_code_cols(ws)

    def _set_zoom(self, ws):
        ws.Activate()
        self.excel.ActiveWindow.Zoom = DIFF_ZOOM_RATIO

    def _format_no_cols(self, ws):
        for f in DIFF_FORMATS['no']:
            ws.Range(f['range']).ColumnWidth = f['width']

    def _format_code_cols(self, ws):
        for f in DIFF_FORMATS['code']:
            ws.Range(f['range']).Font.Name = f['font']

    def _set_home_position(self):
        self.summary_ws.Activate()
        self.summary_ws.Range(HOME_POSITION).Select()

    def _save_book(self):
        self.wb.SaveAs(str(self.output), FileFormat=xlOpenXMLWorkbook)
        print('xlsxへの変換が完了しました。')


if __name__ == '__main__':
    if len(sys.argv) < 3:
        print(f'Usage : {sys.argv[0]} <base> <latest> [<output>]')
        sys.exit(1)

    WinMergeXlsx(*sys.argv[1:4]).generate()
