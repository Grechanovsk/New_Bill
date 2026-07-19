object Form11: TForm11
  Left = 0
  Top = 0
  Caption = 'Form11'
  ClientHeight = 443
  ClientWidth = 860
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Segoe UI'
  Font.Style = []
  Position = poScreenCenter
  OnClose = FormClose
  OnCreate = FormCreate
  TextHeight = 15
  object Label29: TLabel
    Left = 663
    Top = 230
    Width = 34
    Height = 15
    Margins.Left = 4
    Margins.Top = 4
    Margins.Right = 4
    Margins.Bottom = 4
    Caption = #1044#1085#1077#1081':'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = 'Segoe UI'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object Label1: TLabel
    Left = 663
    Top = 39
    Width = 62
    Height = 15
    Margins.Left = 4
    Margins.Top = 4
    Margins.Right = 4
    Margins.Bottom = 4
    Caption = #1053#1072#1095#1080#1085#1072#1103' '#1089':'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = 'Segoe UI'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object Label2: TLabel
    Left = 664
    Top = 103
    Width = 19
    Height = 15
    Margins.Left = 4
    Margins.Top = 4
    Margins.Right = 4
    Margins.Bottom = 4
    Caption = #1055#1086':'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = 'Segoe UI'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object Button1: TButton
    Left = 663
    Top = 409
    Width = 179
    Height = 28
    Caption = #1042#1099#1093#1086#1076
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -15
    Font.Name = 'Segoe UI'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 0
    OnClick = Button1Click
  end
  object MonthCalendar1: TMonthCalendar
    Left = -15
    Top = -4
    Width = 665
    Height = 460
    Date = 45039.000000000000000000
    PopupMenu = PopupMenu1
    TabOrder = 1
    OnClick = MonthCalendar1Click
    OnDblClick = MonthCalendar1DblClick
  end
  object Edit9: TEdit
    Left = 657
    Top = 253
    Width = 179
    Height = 28
    Margins.Left = 4
    Margins.Top = 4
    Margins.Right = 4
    Margins.Bottom = 4
    Alignment = taCenter
    Color = clHotLight
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWhite
    Font.Height = -15
    Font.Name = 'Segoe UI'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 2
  end
  object Edit1: TEdit
    Left = 663
    Top = 62
    Width = 179
    Height = 28
    Margins.Left = 4
    Margins.Top = 4
    Margins.Right = 4
    Margins.Bottom = 4
    Alignment = taCenter
    Color = clHotLight
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWhite
    Font.Height = -15
    Font.Name = 'Segoe UI'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 3
  end
  object Edit2: TEdit
    Left = 664
    Top = 126
    Width = 179
    Height = 28
    Margins.Left = 4
    Margins.Top = 4
    Margins.Right = 4
    Margins.Bottom = 4
    Alignment = taCenter
    Color = clHotLight
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWhite
    Font.Height = -15
    Font.Name = 'Segoe UI'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 4
  end
  object Button2: TButton
    Left = 704
    Top = 184
    Width = 75
    Height = 25
    Caption = #1056#1072#1089#1095#1080#1090#1072#1090#1100
    TabOrder = 5
    OnClick = Button2Click
  end
  object PopupMenu1: TPopupMenu
    OnPopup = PopupMenu1Popup
    Left = 752
    Top = 312
    object N1: TMenuItem
      Caption = #1055#1086#1089#1090#1072#1074#1080#1090#1100'  '#1087#1072#1084#1103#1090#1082#1091'  '#1085#1072' '#1101#1090#1091' '#1076#1072#1090#1091
      OnClick = N1Click
    end
    object N2: TMenuItem
      Caption = #1057#1085#1103#1090#1100'  '#1087#1072#1084#1103#1090#1082#1091'  '#1085#1072' '#1101#1090#1091' '#1076#1072#1090#1091
      OnClick = N2Click
    end
    object N3: TMenuItem
      Caption = #1057#1076#1077#1083#1072#1090#1100' '#1085#1072#1095#1072#1083#1100#1085#1086#1081' '#1076#1072#1090#1086#1081
      OnClick = N3Click
    end
    object N4: TMenuItem
      Caption = #1057#1076#1077#1083#1072#1090#1100' '#1082#1086#1085#1077#1095#1085#1086#1081' '#1076#1072#1090#1086#1081
      OnClick = N4Click
    end
  end
end
