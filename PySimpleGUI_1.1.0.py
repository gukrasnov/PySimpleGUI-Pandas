import pandas
import PySimpleGUI

_combo_1_ = []
_combo_2_ = []

layout = [
            [ PySimpleGUI.Frame
                ( 'Загрузка файлов',
                    [ 
                        [ PySimpleGUI.Text( 'Файл 1', size = ( 6, 1 ), font = ( '', 13 ) ), PySimpleGUI.Input( size = ( None, 1 ), font = ( '', 16 ), enable_events = True, key = '-INPUT-1-', expand_x = True ), PySimpleGUI.FileBrowse( 'Открыть', font = ( '', 13 ) ) ],
                        [ PySimpleGUI.Text( 'Файл 2', size = ( 6, 1 ), font = ( '', 13 ) ), PySimpleGUI.Input( size = ( None, 1 ), font = ( '', 16 ), enable_events = True, key = '-INPUT-2-', expand_x = True ), PySimpleGUI.FileBrowse( 'Открыть', font = ( '', 13 ) ) ],
                    ],
                  expand_x = True
                )
            ],
            [
                PySimpleGUI.Frame
                    ( 'Файл 1',
                        [
                            [ PySimpleGUI.Combo( [ i for i in _combo_1_ ], size = ( None, 1 ), key = '-COMBO-1-1-', font = ( '', 16 ) ) ],
                            [ PySimpleGUI.Combo( [ i for i in _combo_1_ ], size = ( None, 1 ), key = '-COMBO-1-2-', font = ( '', 16 ) ) ]
                        ]
                    ),
                PySimpleGUI.Frame
                    ( ' ',
                        [
                            [ PySimpleGUI.Text( 'Сравнение', size = ( 9, 1 ), font = ( '', 13 ) ) ],
                            [ PySimpleGUI.Button( 'Запись', font = ( '', 13 ), key = '-WRITE-' ) ]
                        ],
                      title_location = 'n', border_width = 0, element_justification = 'Center'
                    ),
                PySimpleGUI.Frame
                    ( 'Файл 2',
                        [
                            [ PySimpleGUI.Combo( [ i for i in _combo_2_ ], size = ( None, 1 ), key = '-COMBO-2-1-', font = ( '', 16 ) ) ],
                            [ PySimpleGUI.Combo( [ i for i in _combo_2_ ], size = ( None, 1 ), key = '-COMBO-2-2-', font = ( '', 16 ) ) ]
                        ]
                    )
            ]
         ]

window = PySimpleGUI.Window( 'Program', layout )

while True:
    event, values = window.read()
    if event == PySimpleGUI.WIN_CLOSED :
        break
    elif event == '-INPUT-1-' :
        path_xlsx_1 = values[ '-INPUT-1-' ]
        work_sheet_1 = pandas.read_excel( path_xlsx_1, sheet_name = 0  )
        df_1 = pandas.DataFrame( work_sheet_1 )
        _combo_1_ = df_1.columns.to_list()
        window[ '-COMBO-1-1-' ].update( values = _combo_1_ )
        window[ '-COMBO-1-2-' ].update( values = _combo_1_ )
    elif event == '-INPUT-2-' :
        path_xlsx_2 = values[ '-INPUT-2-' ]
        work_sheet_2 = pandas.read_excel( path_xlsx_2, sheet_name = 0  )
        df_2 = pandas.DataFrame( work_sheet_2 )
        _combo_2_ = df_2.columns.to_list()
        window[ '-COMBO-2-1-' ].update( values = _combo_2_ )
        window[ '-COMBO-2-2-' ].update( values = _combo_2_ )
    elif event == '-WRITE-' :
        for index_1, ( iRow_1_1, iRow_1_2 ) in enumerate( zip( df_1[ values[ '-COMBO-1-1-' ] ], df_1[ values[ '-COMBO-1-2-' ] ] ) ):
            for index_2, ( iRow_2_1, iRow_2_2 ) in enumerate( zip( df_2[ values[ '-COMBO-2-1-' ] ], df_2[ values[ '-COMBO-2-2-' ] ] ) ):
                if iRow_1_1 == iRow_2_1:
                    df_2[ values[ '-COMBO-2-2-' ] ].iloc[ index_2 ] = iRow_1_2
                    print( index_1, ': =>', iRow_1_1, iRow_1_2, ':::::::::', index_2, ': =>', iRow_2_1, iRow_1_2 )
        with pandas.ExcelWriter( path_xlsx_2, mode = 'a', engine = 'openpyxl', if_sheet_exists = 'overlay' )as writer:
            df_2.to_excel( writer, sheet_name = 'Лист1', index = False, header = False, startcol = 0, startrow = 1 )
    else:
        None

window.close()