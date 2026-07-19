object Form5: TForm5
  Left = 0
  Top = 0
  Caption = '(F5) '#1042#1074#1086#1076' '#1087#1086#1082#1072#1079#1072#1085#1080#1081' '#1089#1095#1077#1090#1095#1080#1082#1072
  ClientHeight = 220
  ClientWidth = 254
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Segoe UI'
  Font.Style = []
  Position = poScreenCenter
  OnActivate = FormActivate
  OnClose = FormClose
  OnCreate = FormCreate
  TextHeight = 15
  object Label1: TLabel
    Left = 19
    Top = 54
    Width = 112
    Height = 15
    Caption = #1055#1086#1082#1072#1079#1072#1085#1080#1103' '#1089#1095#1077#1090#1095#1080#1082#1072
  end
  object Label2: TLabel
    Left = 19
    Top = 2
    Width = 126
    Height = 15
    Caption = #1044#1072#1090#1072' '#1089#1085#1103#1090#1080#1103' '#1087#1086#1082#1072#1079#1072#1085#1080#1081
  end
  object Label3: TLabel
    Left = 19
    Top = 123
    Width = 131
    Height = 15
    Caption = #1056#1072#1079#1085#1080#1094#1072' '#1089' '#1087#1088#1077#1076#1099#1076#1091#1097#1080#1084
  end
  object Button1: TButton
    Left = 161
    Top = 184
    Width = 75
    Height = 27
    Caption = #1047#1072#1082#1088#1099#1090#1100
    TabOrder = 0
    OnClick = Button1Click
  end
  object Edit1: TEdit
    Left = 19
    Top = 75
    Width = 217
    Height = 29
    Alignment = taCenter
    Color = clHotLight
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindow
    Font.Height = -16
    Font.Name = 'Segoe UI'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 1
    OnChange = Edit1Change
    OnKeyPress = Edit1KeyPress
  end
  object Button2: TButton
    Left = 13
    Top = 184
    Width = 75
    Height = 27
    Caption = #1047#1072#1087#1080#1089#1072#1090#1100' '
    TabOrder = 2
    OnClick = Button2Click
  end
  object DateTimePicker1: TDateTimePicker
    Left = 19
    Top = 20
    Width = 217
    Height = 33
    Date = 44756.000000000000000000
    Time = 0.798463252314832100
    Color = clHighlightText
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -19
    Font.Name = 'Segoe UI'
    Font.Style = []
    ParentFont = False
    TabOrder = 3
  end
  object Edit2: TEdit
    Left = 19
    Top = 144
    Width = 217
    Height = 29
    Alignment = taCenter
    Color = clHotLight
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindow
    Font.Height = -16
    Font.Name = 'Segoe UI'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 4
    OnKeyPress = Edit1KeyPress
  end
end
