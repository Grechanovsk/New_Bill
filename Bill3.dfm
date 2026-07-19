object Form3: TForm3
  Left = 0
  Top = 0
  Caption = #1056#1077#1076#1072#1082#1090#1080#1088#1086#1074#1072#1085#1080#1077' '#1073#1072#1079#1099
  ClientHeight = 509
  ClientWidth = 960
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
  object DBGrid2: TDBGrid
    Left = 3
    Top = 25
    Width = 948
    Height = 352
    Margins.Left = 4
    Margins.Top = 4
    Margins.Right = 4
    Margins.Bottom = 4
    DataSource = DataModule1.DataSource_Edit
    Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgTabs, dgRowSelect, dgConfirmDelete, dgCancelOnExit, dgMultiSelect, dgTitleClick, dgTitleHotTrack]
    PopupMenu = Form1.PopupMenu1
    TabOrder = 0
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -12
    TitleFont.Name = 'Segoe UI'
    TitleFont.Style = []
    OnDrawColumnCell = DBGrid2DrawColumnCell
    Columns = <
      item
        Expanded = False
        FieldName = 'id'
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'Sobitie'
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'DenNed'
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'Data'
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'Ostatok'
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'Pil'
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'Travm'
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'Gips'
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'Zapl_inet'
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'Sn_Gips'
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'Den_Gizni'
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'Zam'
        Visible = True
      end>
  end
  object Button1: TButton
    Left = 840
    Top = 477
    Width = 75
    Height = 28
    Caption = #1042#1099#1093#1086#1076
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -15
    Font.Name = 'Segoe UI'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 1
    OnClick = Button1Click
  end
  object Panel4: TPanel
    Left = 0
    Top = 384
    Width = 950
    Height = 87
    Color = clBtnShadow
    ParentBackground = False
    TabOrder = 2
    object Label6: TLabel
      Left = 221
      Top = 1
      Width = 24
      Height = 15
      Margins.Left = 4
      Margins.Top = 4
      Margins.Right = 4
      Margins.Bottom = 4
      Caption = 'Data'
    end
    object Label9: TLabel
      Left = 427
      Top = 1
      Width = 67
      Height = 15
      Margins.Left = 4
      Margins.Top = 4
      Margins.Right = 4
      Margins.Bottom = 4
      Caption = #1055#1080#1083'/'#1053#1077' '#1087#1080#1083
    end
    object Label10: TLabel
      Left = 13
      Top = 1
      Width = 36
      Height = 15
      Margins.Left = 4
      Margins.Top = 4
      Margins.Right = 4
      Margins.Bottom = 4
      Caption = 'Sobitie'
    end
    object Label11: TLabel
      Left = 510
      Top = 2
      Width = 32
      Height = 15
      Margins.Left = 4
      Margins.Top = 4
      Margins.Right = 4
      Margins.Bottom = 4
      Caption = 'Travm'
    end
    object Label12: TLabel
      Left = 575
      Top = 1
      Width = 23
      Height = 15
      Margins.Left = 4
      Margins.Top = 4
      Margins.Right = 4
      Margins.Bottom = 4
      Caption = 'Gips'
    end
    object Label13: TLabel
      Left = 727
      Top = 1
      Width = 41
      Height = 15
      Margins.Left = 4
      Margins.Top = 4
      Margins.Right = 4
      Margins.Bottom = 4
      Caption = 'Sn_Gips'
    end
    object Label14: TLabel
      Left = 648
      Top = 1
      Width = 48
      Height = 15
      Margins.Left = 4
      Margins.Top = 4
      Margins.Right = 4
      Margins.Bottom = 4
      Caption = 'Zapl_inet'
    end
    object Label15: TLabel
      Left = 840
      Top = 2
      Width = 52
      Height = 15
      Margins.Left = 4
      Margins.Top = 4
      Margins.Right = 4
      Margins.Bottom = 4
      Caption = 'Den_Gizni'
    end
    object Label1: TLabel
      Left = 101
      Top = 1
      Width = 43
      Height = 15
      Margins.Left = 4
      Margins.Top = 4
      Margins.Right = 4
      Margins.Bottom = 4
      Caption = 'DenNed'
    end
    object Label2: TLabel
      Left = 319
      Top = 1
      Width = 41
      Height = 15
      Margins.Left = 4
      Margins.Top = 4
      Margins.Right = 4
      Margins.Bottom = 4
      Caption = 'Ostatok'
    end
    object Edit1: TEdit
      Left = 223
      Top = 24
      Width = 93
      Height = 28
      Margins.Left = 4
      Margins.Top = 4
      Margins.Right = 4
      Margins.Bottom = 4
      Alignment = taCenter
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -15
      Font.Name = 'Segoe UI'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 0
      Text = 'Edit1'
    end
    object Edit2: TEdit
      Left = 510
      Top = 25
      Width = 56
      Height = 28
      Margins.Left = 4
      Margins.Top = 4
      Margins.Right = 4
      Margins.Bottom = 4
      Alignment = taCenter
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -15
      Font.Name = 'Segoe UI'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 1
      Text = 'Edit2'
    end
    object Edit3: TEdit
      Left = 575
      Top = 24
      Width = 55
      Height = 28
      Margins.Left = 4
      Margins.Top = 4
      Margins.Right = 4
      Margins.Bottom = 4
      Alignment = taCenter
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -15
      Font.Name = 'Segoe UI'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 2
      Text = 'Edit3'
    end
    object Edit5: TEdit
      Left = 727
      Top = 24
      Width = 95
      Height = 28
      Margins.Left = 4
      Margins.Top = 4
      Margins.Right = 4
      Margins.Bottom = 4
      Alignment = taCenter
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -15
      Font.Name = 'Segoe UI'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 3
      Text = 'Edit5'
    end
    object Edit4: TEdit
      Left = 648
      Top = 24
      Width = 62
      Height = 28
      Margins.Left = 4
      Margins.Top = 4
      Margins.Right = 4
      Margins.Bottom = 4
      Alignment = taCenter
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -15
      Font.Name = 'Segoe UI'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 4
      Text = 'Edit4'
    end
    object Edit6: TEdit
      Left = 840
      Top = 25
      Width = 104
      Height = 28
      Margins.Left = 4
      Margins.Top = 4
      Margins.Right = 4
      Margins.Bottom = 4
      Alignment = taCenter
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -15
      Font.Name = 'Segoe UI'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 5
      Text = 'Edit6'
    end
    object ComboBox2: TComboBox
      Left = 8
      Top = 24
      Width = 89
      Height = 28
      Margins.Left = 4
      Margins.Top = 4
      Margins.Right = 4
      Margins.Bottom = 4
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -15
      Font.Name = 'Segoe UI'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 6
      Items.Strings = (
        #1044#1072#1083#1080' '#1074#1086#1076#1091
        #1055#1089#1088#1083'             '
        #1047#1072#1087#1083#1072#1090#1080#1083' '#1079#1072' '#1080#1085#1077#1090' ')
    end
    object ComboBox1: TComboBox
      Left = 427
      Top = 24
      Width = 75
      Height = 28
      Margins.Left = 4
      Margins.Top = 4
      Margins.Right = 4
      Margins.Bottom = 4
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -15
      Font.Name = 'Segoe UI'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 7
      Items.Strings = (
        #1053#1045' '#1055#1048#1051
        #1055#1086' '#1088#1072#1079#1085#1086#1084#1091
        #1055#1048#1051)
    end
    object ComboBox3: TComboBox
      Left = 100
      Top = 24
      Width = 120
      Height = 28
      Margins.Left = 4
      Margins.Top = 4
      Margins.Right = 4
      Margins.Bottom = 4
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -15
      Font.Name = 'Segoe UI'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 8
      Items.Strings = (
        #1055#1086#1085#1077#1076#1077#1083#1100#1085#1080#1082
        #1042#1090#1086#1088#1085#1080#1082
        #1057#1088#1077#1076#1072
        #1063#1077#1090#1074#1077#1088#1075
        #1055#1103#1090#1085#1080#1094#1072
        #1057#1091#1073#1073#1086#1090#1072
        #1042#1086#1089#1082#1088#1077#1089#1077#1085#1100#1077
        '')
    end
    object CheckBox1: TCheckBox
      Left = 15
      Top = 61
      Width = 77
      Height = 17
      Caption = #1047#1072#1087
      TabOrder = 9
    end
    object CheckBox2: TCheckBox
      Left = 108
      Top = 61
      Width = 77
      Height = 17
      Caption = #1047#1072#1087
      TabOrder = 10
    end
    object CheckBox3: TCheckBox
      Left = 240
      Top = 61
      Width = 77
      Height = 17
      Caption = #1047#1072#1087
      TabOrder = 11
    end
    object CheckBox4: TCheckBox
      Left = 427
      Top = 61
      Width = 77
      Height = 17
      Caption = #1047#1072#1087
      TabOrder = 12
    end
    object Edit7: TEdit
      Left = 322
      Top = 25
      Width = 93
      Height = 28
      Margins.Left = 4
      Margins.Top = 4
      Margins.Right = 4
      Margins.Bottom = 4
      Alignment = taCenter
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -15
      Font.Name = 'Segoe UI'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 13
      Text = 'Edit7'
    end
  end
  object Button2: TButton
    Left = 0
    Top = 477
    Width = 101
    Height = 28
    Caption = #1048#1089#1087#1088#1072#1074#1080#1090#1100
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -15
    Font.Name = 'Segoe UI'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 3
    OnClick = Button2Click
  end
end
