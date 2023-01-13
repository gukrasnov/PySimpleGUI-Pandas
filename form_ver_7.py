import os
import time
import pandas
import datetime
import PySimpleGUI

start_time = time.monotonic()

_file_1_combo_1_ = []
_file_1_combo_2_ = []
_file_2_combo_1_ = []
_file_2_combo_2_ = []

layot = [
          [
            PySimpleGUI.T( '1 файл' ),
            PySimpleGUI.In( enable_events = True, key = '_file_1_' ),
            PySimpleGUI.FileBrowse( button_text = 'Выбрать 1 файл', target = '_file_1_' )
          ],
          [
            PySimpleGUI.T( '2 файл' ),
            PySimpleGUI.In( enable_events = True, key = '_file_2_' ),
            PySimpleGUI.FileBrowse( button_text = 'Выбрать 2 файл', target = '_file_2_' )
          ],
          [
            PySimpleGUI.Column( [ [ PySimpleGUI.Frame( '1 файл', [ [ PySimpleGUI.Combo( [ i for i in _file_1_combo_1_ ], size = ( 21 ), key = '_file_1_combo_1_' ) ], [ PySimpleGUI.Combo( [ i for i in _file_1_combo_2_ ], size = ( 21 ), key = '_file_1_combo_2_' ) ] ] ) ] ] ),
            PySimpleGUI.Column( [ [ PySimpleGUI.T( 'Сравнение' ) ], [ PySimpleGUI.T( 'Замена' ) ] ], element_justification = 'center' ),
            PySimpleGUI.Column( [ [ PySimpleGUI.Frame( '2 файл', [ [ PySimpleGUI.Combo( [ i for i in _file_2_combo_1_ ], size = ( 21 ), key = '_file_2_combo_1_' ) ], [ PySimpleGUI.Combo( [ i for i in _file_2_combo_2_ ], size = ( 21 ), key = '_file_2_combo_2_' ) ] ] ) ] ] )
          ],
          [
            PySimpleGUI.Column( [ [ PySimpleGUI.Button( 'Заменить', key = '_write_' ) ] ] )
          ]
        ]

window = PySimpleGUI.Window( 'Excel', layot, element_justification = 'center', resizable = True, grab_anywhere = True )

while True:
  event, values = window.read()
  if event == PySimpleGUI.WIN_CLOSED or event == 'Cancel':
    break
  elif event == '_file_1_' :
    _file_1_work_sheet_1_ = pandas.read_excel( values[ '_file_1_' ], sheet_name = 'Лист1' )
    _file_1_work_sheet_2_ = pandas.read_excel( values[ '_file_1_' ], sheet_name = 'Лист1' )
    _file_1_df_1_ = pandas.DataFrame( _file_1_work_sheet_1_ )
    _file_1_df_2_ = pandas.DataFrame( _file_1_work_sheet_2_ )
    _file_1_combo_1_ = _file_1_df_1_.columns.to_list()
    _file_1_combo_2_ = _file_1_df_2_.columns.to_list()
    window[ '_file_1_combo_1_' ].update( values = _file_1_combo_1_ )
    window[ '_file_1_combo_2_' ].update( values = _file_1_combo_2_ )
  elif event == '_file_2_' :
    _file_2_work_sheet_1_ = pandas.read_excel( values[ '_file_2_' ], sheet_name = 'Лист1' )
    _file_2_work_sheet_2_ = pandas.read_excel( values[ '_file_2_' ], sheet_name = 'Лист1' )
    _file_2_df_1_ = pandas.DataFrame( _file_2_work_sheet_1_ )
    _file_2_df_2_ = pandas.DataFrame( _file_2_work_sheet_2_ )
    _file_2_combo_1_ = _file_2_df_1_.columns.to_list()
    _file_2_combo_2_ = _file_2_df_2_.columns.to_list()
    window[ '_file_2_combo_1_' ].update( values = _file_2_combo_1_ )
    window[ '_file_2_combo_2_' ].update( values = _file_2_combo_2_ )
  elif event == '_write_':
    for index_1, ( iRow_1_1, iRow_1_2 ) in enumerate( zip( _file_1_df_1_[ values[ '_file_1_combo_1_' ] ], _file_1_df_1_[ values[ '_file_1_combo_2_' ] ] ) ):
      for index_2, ( iRow_2_1, iRow_2_2 ) in enumerate( zip( _file_2_df_1_[ values[ '_file_2_combo_1_' ] ], _file_2_df_2_[ values[ '_file_2_combo_2_' ] ] ) ):
        if iRow_1_1 == iRow_2_1:
            _file_1_df_2_[ values[ '_file_2_combo_2_' ] ].iloc[ index_2 ] = iRow_1_2
            print( index_1, ': =>', iRow_1_1, iRow_1_2, ':::::::::', index_2, ': =>', iRow_2_1, iRow_1_2 )
    _file_1_df_2_.to_excel( '3333333.xlsx', sheet_name = 'Лист1', index = False, startcol = 0, startrow = 1 )
  else:
    print( 'Ты вошёл: ', values[ 0 ] )

window.close()

end_time = time.monotonic()
print( 'Время выполнения кода: ', datetime.timedelta( seconds = end_time - start_time ) )