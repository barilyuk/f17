object ZvitF: TZvitF
  Left = 245
  Top = 112
  AutoSize = True
  BorderStyle = bsDialog
  Caption = #1047#1074#1110#1090
  ClientHeight = 97
  ClientWidth = 393
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object GroupBox1: TGroupBox
    Left = 0
    Top = 0
    Width = 241
    Height = 97
    Caption = #1042#1080#1079#1085#1072#1095#1090#1077' '#1110#1085#1090#1077#1088#1074#1072#1083' '#1079#1074#1110#1090#1091
    TabOrder = 0
    object Label1: TLabel
      Left = 16
      Top = 26
      Width = 93
      Height = 13
      Caption = #1055#1086#1095#1072#1090#1086#1082' '#1110#1085#1090#1077#1088#1074#1072#1083#1091
    end
    object Label2: TLabel
      Left = 16
      Top = 60
      Width = 84
      Height = 13
      Caption = #1050#1110#1085#1077#1094#1100' '#1110#1085#1090#1077#1088#1074#1072#1083#1091
    end
    object dtpBeg: TDateTimePicker
      Left = 128
      Top = 24
      Width = 97
      Height = 21
      CalAlignment = dtaLeft
      Date = 37045.6667814815
      Time = 37045.6667814815
      DateFormat = dfShort
      DateMode = dmComboBox
      Kind = dtkDate
      MinDate = 36526
      ParseInput = False
      TabOrder = 0
    end
    object dtpEnd: TDateTimePicker
      Left = 128
      Top = 56
      Width = 97
      Height = 21
      CalAlignment = dtaLeft
      Date = 37045.6668131944
      Time = 37045.6668131944
      DateFormat = dfShort
      DateMode = dmComboBox
      Kind = dtkDate
      MinDate = 36526
      ParseInput = False
      TabOrder = 1
    end
  end
  object GroupBox2: TGroupBox
    Left = 248
    Top = 0
    Width = 145
    Height = 97
    Caption = #1057#1092#1086#1088#1084#1091#1074#1072#1090#1080' '#1079#1074#1110#1090
    TabOrder = 1
    object BitBtn1: TBitBtn
      Left = 24
      Top = 24
      Width = 100
      Height = 25
      Caption = #1057#1092#1086#1088#1084#1091#1074#1072#1090#1080
      Default = True
      TabOrder = 0
      OnClick = BitBtn1Click
      Glyph.Data = {
        DE010000424DDE01000000000000760000002800000024000000120000000100
        0400000000006801000000000000000000001000000000000000000000000000
        80000080000000808000800000008000800080800000C0C0C000808080000000
        FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333333333
        3333333333333333333333330000333333333333333333333333F33333333333
        00003333344333333333333333388F3333333333000033334224333333333333
        338338F3333333330000333422224333333333333833338F3333333300003342
        222224333333333383333338F3333333000034222A22224333333338F338F333
        8F33333300003222A3A2224333333338F3838F338F33333300003A2A333A2224
        33333338F83338F338F33333000033A33333A222433333338333338F338F3333
        0000333333333A222433333333333338F338F33300003333333333A222433333
        333333338F338F33000033333333333A222433333333333338F338F300003333
        33333333A222433333333333338F338F00003333333333333A22433333333333
        3338F38F000033333333333333A223333333333333338F830000333333333333
        333A333333333333333338330000333333333333333333333333333333333333
        0000}
      NumGlyphs = 2
    end
    object BitBtn2: TBitBtn
      Left = 24
      Top = 56
      Width = 100
      Height = 25
      Cancel = True
      Caption = #1055#1086#1074#1077#1088#1085#1091#1090#1080#1089#1103
      TabOrder = 1
      OnClick = BitBtn2Click
      Glyph.Data = {
        DE010000424DDE01000000000000760000002800000024000000120000000100
        0400000000006801000000000000000000001000000000000000000000000000
        80000080000000808000800000008000800080800000C0C0C000808080000000
        FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333333333
        333333333333333333333333000033338833333333333333333F333333333333
        0000333911833333983333333388F333333F3333000033391118333911833333
        38F38F333F88F33300003339111183911118333338F338F3F8338F3300003333
        911118111118333338F3338F833338F3000033333911111111833333338F3338
        3333F8330000333333911111183333333338F333333F83330000333333311111
        8333333333338F3333383333000033333339111183333333333338F333833333
        00003333339111118333333333333833338F3333000033333911181118333333
        33338333338F333300003333911183911183333333383338F338F33300003333
        9118333911183333338F33838F338F33000033333913333391113333338FF833
        38F338F300003333333333333919333333388333338FFF830000333333333333
        3333333333333333333888330000333333333333333333333333333333333333
        0000}
      NumGlyphs = 2
    end
  end
  object ExApp: TExcelApplication
    AutoConnect = True
    ConnectKind = ckRunningOrNew
    AutoQuit = True
    Left = 232
  end
  object Exbook: TExcelWorkbook
    AutoConnect = False
    ConnectKind = ckAttachToInterface
    Left = 256
  end
end
