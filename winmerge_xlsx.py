import sys
import os
import shutil
from pathlib import Path
import subprocess
import json
import win32com.client

WINMERGE_EXE = r'C:\Program Files\WinMerge\WinMergeU.exe'  # WinMergeへのパス
WINMERGE_OPTIONS = [
    '/minimize',                             # ウィンドウ最小化で起動
    '/noninteractive',                       # レポート出力後に終了
    '/cfg',
    'Settings/DirViewExpandSubdirs=1',       # 自動的にサブフォルダーを展開する
    '/cfg',
    'Settings/ViewLineNumbers=1',            # 行番号を表示する
    '/cfg',
    'Settings/WordDifferenceTextColor=255',  # 差分単語を赤色で表示
    '/cfg',
    'ReportFiles/ReportType=2',              # シンプルなHTML形式
    '/cfg',
    'ReportFiles/IncludeFileCmpReport=1',    # ファイル比較レポートを含める
    '/r',                                    # すべてのサブフォルダ内のすべてのファイルを比較
    '/u',                                    # 最近使用した項目リストに追加しない
    '/or',                                   # レポートを出力
]

xlUp = -4162
xlToLeft = -4159
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
    """WinMergeの差分レポートをエクセルに出力
    """
    def __init__(self, base, latest, output='./output.xlsx'):
        self.base = Path(base).absolute()
        self.latest = Path(latest).absolute()
        self.output = Path(output).absolute()

        parent = str(self.output.parent)
        stem = str(self.output.stem)
        self.output_html = Path(parent + '/' + stem + '.html')
        self.output_html_files = Path(parent + '/' + stem + '.files')

        self.setting_json = './setting.json'
        if os.path.exists(self.setting_json):
            with open(self.setting_json, 'r') as f:
                json_load = json.load(f)
            if 'WINMERGE_EXE' in json_load:
                global WINMERGE_EXE
                WINMERGE_EXE = json_load['WINMERGE_EXE']
            if 'WINMERGE_OPTIONS' in json_load:
                global WINMERGE_OPTIONS
                WINMERGE_OPTIONS = json_load['WINMERGE_OPTIONS']
            if 'DIFF_FORMATS' in json_load:
                global DIFF_FORMATS
                DIFF_FORMATS = json_load['DIFF_FORMATS']

        self.sheet_memo = {}
        self.sheet_count = {}

    def generate(self):
        """レポート生成
        """
        self._setup()
        self._generate_html_by_winmerge()
        self._convert_html_to_xlsx()

    def _setup(self):
        """準備
        """
        self._setup_excel_application()
        self._setup_output_files()

    def _setup_excel_application(self):
        """エクセルアプリケーションの準備
        """
        try:
            if win32com.client.GetObject(Class='Excel.Application'):
                self.__message_and_exit('Excelを閉じて下さい。')
        except win32com.client.pywintypes.com_error:
            pass

    def _setup_output_files(self):
        """出力するファイルの準備
        """
        # htmlレポート
        if (os.path.exists(self.output_html)):
            try:
                os.remove(self.output_html)
            except PermissionError:
                message = str(self.output_html) + 'へのアクセス権がありません。'
                self.__message_and_exit(message)
        # WinMergeの中間ファイル
        if (os.path.isdir(self.output_html_files)):
            try:
                shutil.rmtree(self.output_html_files)
            except PermissionError:
                message = str(self.output_html_files) + 'へのアクセス権がありません。'
                self.__message_and_exit(message)
        # エクセルレポートファイル
        if (os.path.exists(self.output)):
            try:
                os.remove(self.output)
            except PermissionError:
                message = str(self.output) + 'へのアクセス権がありません。'
                self.__message_and_exit(message)

    def __message_and_exit(self, message):
        """メッセージを表示して終了
        """
        print('\nError : ' + message)
        sys.exit(-1)

    def _generate_html_by_winmerge(self):
        """WinMergeにてhtmlレポート生成
        """
        print("\n[generate html by winmerge]")

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
        """htmlレポートをエクセルファイルに変換する
        """
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
        """ブックを開く
        """
        self.excel = win32com.client.Dispatch('Excel.Application')
        self.wb = self.excel.Workbooks.Open(self.output_html)
        self.summary_ws = self.wb.Worksheets(SUMMARY_WS_NUM)

    def _format_summary_sheet(self):
        """一覧シートの書式調整
        """
        print("\n[format summary sheet]")

        ws = self.summary_ws
        end_row = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        # 同名シート有無の確認
        print('- check same filename ...', end='', flush=True)
        for row in range(SUMMARY_START_ROW, end_row+1):
            name_cell = ws.Range(SUMMARY_NAME_COL + str(row))
            if not name_cell.Value:
                break
            if name_cell.Hyperlinks.Count > 0:
                sheet_name = name_cell.Value.lower()
                if sheet_name in self.sheet_memo:
                    self.sheet_memo[sheet_name] += 1
                else:
                    self.sheet_memo[sheet_name] = 1
        print(' done')

        # htmlファイルとハイパーリンクの設定
        for row in range(SUMMARY_START_ROW, end_row+1):
            name_cell = ws.Range(SUMMARY_NAME_COL + str(row))
            if not name_cell.Value:
                break

            if name_cell.Hyperlinks.Count > 0:
                sname_src = sname_dst = name_cell.Value
                sname_src_lower = sname_src.lower()
                if sname_src_lower in self.sheet_count:
                    self.sheet_count[sname_src_lower] += 1
                else:
                    self.sheet_count[sname_src_lower] = 1

                if self.sheet_memo[sname_src_lower] >= 2:
                    sname_dst += f'_{self.sheet_count[sname_src_lower]}'

                self._change_hyperlink(name_cell, sname_dst)
                folder_cell = ws.Range(SUMMARY_FOLDER_COL + str(row)).Value
                self._rename_html_files(sname_src, sname_dst, folder_cell)

    def _change_hyperlink(self, name_cell, name_dst):
        """ハイパーリンクの修正
        """
        for hl in name_cell.Hyperlinks:
            hl.Address = ''
            hl.SubAddress = name_dst + '!' + HOME_POSITION
            hl.TextToDisplay = name_dst

    def _rename_html_files(self, name_src, name_dst, folder):
        """htmlレポートのリネーム
        """
        sheet_name = folder.replace('\\', '_') + '_' + name_src if folder else name_src  # noqa: E501
        src = f'{self.output_html_files}/{sheet_name}.html'
        dst = f'{self.output_html_files}/{name_dst}.html'
        if sheet_name != name_dst:
            os.rename(src, dst)
            print('- ' + sheet_name + ' ---> ' + name_dst)

    def _copy_html_files(self):
        """htmlレポートをエクセルにコピー
        """
        print("\n[copy html files]")

        count = 1
        for html in self.output_html_files.glob('**/*.html'):
            try:
                print(f'- {html.name} ...', end='', flush=True)
                diff_wb = self.excel.Workbooks.Open(html)
                diff_ws = diff_wb.Worksheets(1)
                diff_ws.Copy(Before=None, After=self.wb.Worksheets(count))
                diff_wb.Close()
                count += 1
                print(' done')
            except win32com.client.pywintypes.com_error as e:
                print(' skipped *** unknown error ***')
                print(f'\n{e}\n')

    def _format_diff_sheets(self):
        """差分シートの書式調整
        """
        print("\n[format diff sheets]")

        for i in range(DIFF_START_ROW, self.wb.Worksheets.Count+1):
            ws = self.wb.Worksheets(i)
            print(f'- {ws.Name} ...', end='', flush=True)
            self._set_zoom(ws)
            self._freeze_panes(ws)
            self._remove_hyperlink_from_no(ws)
            self._set_format(ws)
            self._set_autofilter(ws)
            self._set_home_position(ws)
            print(' done')

    def _set_zoom(self, ws):
        """拡大率を設定する
        """
        ws.Activate()
        self.excel.ActiveWindow.Zoom = DIFF_ZOOM_RATIO

    def _freeze_panes(self, ws):
        """ウィンドウ枠を固定する
        """
        ws.Activate()
        ws.Range('A' + str(DIFF_START_ROW)).Select()
        self.excel.ActiveWindow.FreezePanes = True

    def _remove_hyperlink_from_no(self, ws):
        """行番号のハイパーリンクを削除する
        """
        for f in DIFF_FORMATS['no']:
            end_row = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            r1 = f['col'] + ':' + f['col']
            r2 = f['col'] + str(DIFF_START_ROW) + ':' + f['col'] + str(end_row)
            ws.Range(r1).Hyperlinks.Delete()
            ws.Range(r2).Interior.Color = int('F0F0F0', 16)
            ws.Range(r2).Font.Size = 12

    def _set_format(self, ws):
        """差分シートの書式調整
        """
        for key in DIFF_FORMATS.keys():
            for f in DIFF_FORMATS[key]:
                r = f['col'] + ':' + f['col']
                ws_range = ws.Range(r)
                if 'width' in f:
                    ws_range.ColumnWidth = f['width']
                if 'font' in f:
                    ws_range.Font.Name = f['font']
                if 'header' in f:
                    self._set_extra_table(ws, f)

    def _set_extra_table(self, ws, f):
        """表を追加する
        """
        end_row = self._get_diff_end_row(ws)

        # 表の見出し
        ws_range = ws.Range(f['col'] + '1')
        ws_range.Value = f['header']
        ws_range.VerticalAlignment = xlCenter
        ws_range.HorizontalAlignment = xlCenter
        ws_range.Interior.Color = int('CCFFCC', 16)

        # 表の罫線
        ws_range = ws.Range(f['col'] + '1:' + f['col'] + str(end_row))
        ws_range.Borders.Color = int('000000', 16)
        ws_range.Borders.LineStyle = xlContinuous

        # 表の中身を一旦"-"で埋める
        ws_range = ws.Range(f['col'] + '2:' + f['col'] + str(end_row))
        ws_range.Value = '-'
        ws_range.Interior.Color = int('E0E0E0', 16)

        # 差分がある箇所を空欄に変更する
        code_col = DIFF_FORMATS['code'][0]['col']
        target = code_col + str(DIFF_START_ROW) + ':' + code_col + str(end_row)
        row = DIFF_START_ROW
        group = 0
        for cell in ws.Range(target):
            if cell.Interior.Color != int('FFFFFF', 16):
                group += 1
            else:
                if group:
                    ws_range = ws.Range(f['col'] + str(row-group) + ':' + f['col'] + str(row-1))  # noqa: E501
                    ws_range.Value = ''
                    ws_range.Interior.Color = int('FFFFFF', 16)
                    group = 0
            row += 1
        if group:
            ws_range = ws.Range(f['col'] + str(end_row-group+1) + ':' + f['col'] + str(end_row))  # noqa: E501
            ws_range.Value = ''
            ws_range.Interior.Color = int('FFFFFF', 16)

    def _set_autofilter(self, ws):
        """オートフィルタを設定する
        """
        start = 'E1'
        end = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Address.split('$')[1] + '1'  # noqa: E501
        if end == start:
            end = 'F1'
        ws.Range(start, end).AutoFilter()

    def _set_home_position(self, ws):
        """ホームポジションを設定する
        """
        ws.Activate()
        ws.Range(HOME_POSITION).Select()

    def _save_book(self):
        """ブックを保存する
        """
        self.wb.SaveAs(str(self.output), FileFormat=xlOpenXMLWorkbook)
        print('\nxlsxへの変換が完了しました。')

    def _get_diff_end_row(self, ws):
        """差分シートの最終行の番号を取得する
        """
        left_row = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
        right_row = ws.Cells(ws.Rows.Count, 4).End(xlUp).Row
        return max(left_row, right_row) + 1


if __name__ == '__main__':
    if len(sys.argv) < 3:
        print(f'Usage : {sys.argv[0]} <base> <latest> [<output>]')
        sys.exit(1)

    import time

    start = time.perf_counter()
    WinMergeXlsx(*sys.argv[1:4]).generate()
    end = time.perf_counter()

    print(f'elp = {end-start:.3f}s')
