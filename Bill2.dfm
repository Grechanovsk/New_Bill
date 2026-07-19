object Form2: TForm2
  Left = 0
  Top = 0
  Caption = #1054#1087#1083#1072#1090#1072
  ClientHeight = 143
  ClientWidth = 332
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
  object Label1: TLabel
    Left = 44
    Top = 24
    Width = 136
    Height = 23
    Caption = #1044#1072#1090#1072' '#1087#1083#1072#1090#1105#1078#1072':    :'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -17
    Font.Name = 'Segoe UI'
    Font.Style = []
    ParentFont = False
  end
  object Label2: TLabel
    Left = 29
    Top = 74
    Width = 151
    Height = 23
    Caption = #1057#1091#1084#1084#1072' '#1087#1083#1072#1090#1105#1078#1072':    :'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -17
    Font.Name = 'Segoe UI'
    Font.Style = []
    ParentFont = False
  end
  object DateTimePicker1: TDateTimePicker
    Left = 185
    Top = 24
    Width = 104
    Height = 23
    Date = 44543.000000000000000000
    Time = 0.516766041670052800
    TabOrder = 0
  end
  object SpinEdit1: TSpinEdit
    Left = 185
    Top = 66
    Width = 76
    Height = 38
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -20
    Font.Name = 'Segoe UI'
    Font.Style = [fsBold]
    MaxValue = 0
    MinValue = 0
    ParentFont = False
    TabOrder = 1
    Value = 0
  end
  object Button1: TButton
    Left = 119
    Top = 110
    Width = 75
    Height = 25
    Caption = 'Ok'
    TabOrder = 2
    OnClick = Button1Click
  end
  object Button2: TButton
    Left = 264
    Top = 58
    Width = 39
    Height = 25
    Caption = '+10'
    TabOrder = 3
    OnClick = Button2Click
  end
  object Button3: TButton
    Left = 264
    Top = 92
    Width = 39
    Height = 25
    Caption = '+100'
    TabOrder = 4
    OnClick = Button3Click
  end
end
