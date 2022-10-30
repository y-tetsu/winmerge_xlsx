import sys
import os
import shutil
from pathlib import Path
import subprocess
import win32com.client

WINMERGE_EXE = r'C:\Program Files\WinMerge\WinMergeU.exe'  # WinMergeへのパス
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
xlCenter = -4108
xlContinuous = 1

SUMMARY_WS_NUM = 1        # 一覧シートのワークシート番号
SUMMARY_START_ROW = 6     # 一覧シートの表の開始行
SUMMARY_NAME_COL = 'A'    # 一覧シートの表の名前列
SUMMARY_FOLDER_COL = 'B'  # 一覧シートの表のフォルダー列

HOME_POSITION = 'A1'  # ホームポジション

DIFF_START_ROW = 2                                            # 差分シートの開始行
DIFF_ZOOM_RATIO = 85                                          # 差分シートのズームの倍率
DIFF_FORMATS = {                                              # 差分シートの書式設定
    'no': [                                                   # 行番号列
        {'col': 'A', 'width': 5},                             # 左側
        {'col': 'C', 'width': 5},                             # 右側
    ],
    'code': [                                                 # ソースコード列
        {'col': 'B', 'width': 100, 'font': 'ＭＳ ゴシック'},  # 左側
        {'col': 'D', 'width': 100, 'font': 'ＭＳ ゴシック'},  # 右側
    ],
    'extra': [                                                # 追加列
        {'col': 'E', 'width': 60, 'header': 'コメント'},
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
                self.__message_and_exit('Excelを閉じて下さい。')
        except win32com.client.pywintypes.com_error:
            pass

    def _setup_output_files(self):
        if (os.path.exists(self.output_html)):
            try:
                os.remove(self.output_html)
            except PermissionError:
                message = str(self.output_html) + 'へのアクセス権がありません。'
                self.__message_and_exit(message)
        if (os.path.isdir(self.output_html_files)):
            try:
                shutil.rmtree(self.output_html_files)
            except PermissionError:
                message = str(self.output_html_files) + 'へのアクセス権がありません。'
                self.__message_and_exit(message)
        if (os.path.exists(self.output)):
            try:
                os.remove(self.output)
            except PermissionError:
                message = str(self.output) + 'へのアクセス権がありません。'
                self.__message_and_exit(message)

    def __message_and_exit(self, message):
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
            self._set_home_position(self.summary_ws)
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
                    self._rename_html_files(name_cell.Value, folder_cell.Value)

    def _change_hyperlink(self, name_cell):
        for hl in name_cell.Hyperlinks:
            hl.Address = ''
            hl.SubAddress = name_cell.Value + '!' + HOME_POSITION

    def _rename_html_files(self, name, folder):
        sheet_name = folder.replace('\\', '_') + '_' + name
        src = f'{self.output_html_files}/{sheet_name}.html'
        dst = f'{self.output_html_files}/{name}.html'
        os.rename(src, dst)
        print(sheet_name + ' ---> ' + name)

    def _copy_html_files(self):
        g = self.output_html_files.glob('**/*.html')
        for count, html in enumerate(g, 1):
            diff_wb = self.excel.Workbooks.Open(html)
            diff_ws = diff_wb.Worksheets(1)
            diff_ws.Copy(Before=None, After=self.wb.Worksheets(count))

    def _format_diff_sheets(self):
        for i in range(DIFF_START_ROW, self.wb.Worksheets.Count+1):
            ws = self.wb.Worksheets(i)
            self._set_zoom(ws)
            self._freeze_panes(ws)
            self._remove_hyperlink_from_no(ws)
            self._set_format(ws)
            self._set_home_position(ws)

    def _set_zoom(self, ws):
        ws.Activate()
        self.excel.ActiveWindow.Zoom = DIFF_ZOOM_RATIO

    def _freeze_panes(self, ws):
        ws.Activate()
        ws.Range('A' + str(DIFF_START_ROW)).Select()
        self.excel.ActiveWindow.FreezePanes = True

    def _remove_hyperlink_from_no(self, ws):
        for f in DIFF_FORMATS['no']:
            end_row = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            r1 = f['col'] + ':' + f['col']
            r2 = f['col'] + str(DIFF_START_ROW) + ':' + f['col'] + str(end_row)
            ws.Range(r1).Hyperlinks.Delete()
            ws.Range(r2).Interior.Color = int('F0F0F0', 16)
            ws.Range(r2).Font.Size = 12

    def _set_format(self, ws):
        for key in DIFF_FORMATS.keys():
            for f in DIFF_FORMATS[key]:
                r = f['col'] + ':' + f['col']
                if 'width' in f:
                    ws.Range(r).ColumnWidth = f['width']
                if 'font' in f:
                    ws.Range(r).Font.Name = f['font']
                if 'header' in f:
                    self._set_extra_table(ws, f)

    def _set_extra_table(self, ws, f):
        end_row = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        r1 = f['col'] + '1'
        r2 = f['col'] + '1:' + f['col'] + str(end_row)
        ws.Range(r1).Value = f['header']
        ws.Range(r1).VerticalAlignment = xlCenter
        ws.Range(r1).HorizontalAlignment = xlCenter
        ws.Range(r1).Interior.Color = int('CCFFCC', 16)
        ws.Range(r2).Borders.Color = int('000000', 16)
        ws.Range(r2).Borders.LineStyle = xlContinuous
        for i in range(DIFF_START_ROW, end_row+1):
            code_col = DIFF_FORMATS['code'][0]['col']
            code_color = ws.Range(code_col + str(i)).Interior.Color
            if code_color == int('FFFFFF', 16):
                r = f['col'] + str(i)
                ws.Range(r).Value = '-'
                ws.Range(r).Interior.Color = int('E0E0E0', 16)

    def _set_home_position(self, ws):
        ws.Activate()
        ws.Range(HOME_POSITION).Select()

    def _save_book(self):
        self.wb.SaveAs(str(self.output), FileFormat=xlOpenXMLWorkbook)
        print('xlsxへの変換が完了しました。')


if __name__ == '__main__':
    if len(sys.argv) < 3:
        print(f'Usage : {sys.argv[0]} <base> <latest> [<output>]')
        sys.exit(1)

    WinMergeXlsx(*sys.argv[1:4]).generate()
