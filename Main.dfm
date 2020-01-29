object Form1: TForm1
  Left = 273
  Top = 268
  Width = 404
  Height = 266
  Caption = #1047#1074#1110#1090#1080
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 16
    Top = 16
    Width = 237
    Height = 20
    Caption = #1042#1080#1073#1077#1088#1110#1090#1100' '#1085#1077#1086#1073#1093#1110#1076#1085#1080#1081' '#1087#1077#1088#1110#1086#1076
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object Label2: TLabel
    Left = 16
    Top = 40
    Width = 10
    Height = 20
    Caption = #1079
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object Label3: TLabel
    Left = 216
    Top = 40
    Width = 21
    Height = 20
    Caption = #1087#1086
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object BitBtn1: TBitBtn
    Left = 16
    Top = 72
    Width = 361
    Height = 25
    Caption = #1042#1080#1082#1086#1085#1072#1090#1080' '#1060'.17'
    TabOrder = 0
    OnClick = BitBtn1Click
    Kind = bkOK
  end
  object date1: TDateTimePicker
    Left = 48
    Top = 40
    Width = 129
    Height = 21
    CalAlignment = dtaLeft
    Date = 43101.4792852778
    Time = 43101.4792852778
    DateFormat = dfShort
    DateMode = dmComboBox
    Kind = dtkDate
    ParseInput = False
    TabOrder = 1
  end
  object date2: TDateTimePicker
    Left = 248
    Top = 40
    Width = 129
    Height = 21
    CalAlignment = dtaLeft
    Date = 43132.4794228588
    Time = 43132.4794228588
    DateFormat = dfShort
    DateMode = dmComboBox
    Kind = dtkDate
    ParseInput = False
    TabOrder = 2
  end
  object BitBtn2: TBitBtn
    Left = 16
    Top = 104
    Width = 361
    Height = 25
    Caption = #1047#1074#1110#1090' '#1087#1086' '#1079#1072#1084#1086#1074#1083#1077#1085#1085#1103#1084' ('#1087#1110#1074#1087#1072#1088#1080') '#1087#1086' '#1076#1072#1090#1110' '#1087#1086#1076#1072#1085#1085#1103
    TabOrder = 3
    OnClick = BitBtn2Click
    Kind = bkAll
  end
  object BitBtn3: TBitBtn
    Left = 16
    Top = 136
    Width = 361
    Height = 25
    Caption = #1047#1074#1110#1090' '#1087#1086' '#1079#1072#1084#1086#1074#1083#1077#1085#1085#1103#1084' ('#1087#1110#1074#1087#1072#1088#1080') '#1087#1086' '#1076#1072#1090#1110' '#1074#1080#1076#1072#1095#1110
    TabOrder = 4
    OnClick = BitBtn3Click
    Kind = bkAll
  end
  object BitBtn4: TBitBtn
    Left = 16
    Top = 168
    Width = 361
    Height = 25
    Caption = #1057#1087#1080#1089#1086#1082' '#1074#1080#1082#1086#1088#1080#1089#1090#1072#1085#1080#1093' '#1084#1072#1090#1077#1088#1110#1072#1083#1110#1074' ('#1079#1072' '#1096#1080#1092#1088#1072#1084#1080')'
    TabOrder = 5
    OnClick = BitBtn4Click
    Kind = bkAll
  end
  object BitBtn5: TBitBtn
    Left = 16
    Top = 200
    Width = 361
    Height = 25
    Caption = #1057#1087#1080#1089#1086#1082' '#1074#1080#1082#1086#1088#1080#1089#1090#1072#1085#1080#1093' '#1084#1072#1090#1077#1088#1110#1072#1083#1110#1074' ('#1079#1072' '#1072#1088#1090#1080#1082#1091#1083#1072#1084#1080')'
    TabOrder = 6
    OnClick = BitBtn5Click
    Kind = bkAll
  end
  object DataSource1: TDataSource
    DataSet = Query1
    Left = 320
  end
  object ExApp: TExcelApplication
    AutoConnect = True
    ConnectKind = ckRunningOrNew
    AutoQuit = True
    Left = 352
  end
  object Exbook: TExcelWorkbook
    AutoConnect = False
    ConnectKind = ckAttachToInterface
    Left = 192
  end
  object Query1: TQuery
    DatabaseName = 'DataBase1'
    SQL.Strings = (
      'DROP TABLE t11_temp; '
      'SELECT     N_KAR, COUNT(N_KAR) AS Perv'
      'INTO            T11_temp'
      'FROM          RS_T11'
      'WHERE      (N_ZAV = 3)'
      'GROUP BY N_KAR')
    Left = 280
  end
  object Database1: TDatabase
    AliasName = 'SQL_UKP'
    DatabaseName = 'DataBase1'
    LoginPrompt = False
    SessionName = 'Default'
    Left = 240
  end
end
