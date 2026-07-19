object DataModule1: TDataModule1
  Height = 633
  Width = 640
  object FDConnection_osn: TFDConnection
    Params.Strings = (
      'DriverID=SQLite')
    LoginPrompt = False
    Left = 39
    Top = 74
  end
  object FDQuery_old_1: TFDQuery
    Connection = FDConnection_osn
    Left = 148
    Top = 10
  end
  object FDQuery_osn: TFDQuery
    Connection = FDConnection_osn
    SQL.Strings = (
      'select * from Bill_New ')
    Left = 146
    Top = 74
  end
  object FDQuery_Svet: TFDQuery
    Connection = FDConnection_osn
    SQL.Strings = (
      'select * from Svet')
    Left = 141
    Top = 246
  end
  object DataSource_osn: TDataSource
    DataSet = FDQuery_osn
    Left = 240
    Top = 71
  end
  object FDQuery_opl: TFDQuery
    Connection = FDConnection_osn
    SQL.Strings = (
      'SELECT * FROM Opl_table')
    Left = 143
    Top = 138
  end
  object DataSource_opl: TDataSource
    DataSet = FDQuery_opl
    Left = 240
    Top = 136
  end
  object FDQuery_excel: TFDQuery
    Connection = FDConnection_osn
    SQL.Strings = (
      'select * from Bill_New')
    Left = 143
    Top = 187
  end
  object DataSource_excel: TDataSource
    DataSet = FDQuery_excel
    Left = 239
    Top = 187
  end
  object DataSource_Svet: TDataSource
    DataSet = FDQuery_Svet
    Left = 238
    Top = 247
  end
  object FDQuery_Kvartira: TFDQuery
    Connection = FDConnection_osn
    SQL.Strings = (
      'SELECT * FROM Kvartira ')
    Left = 143
    Top = 302
  end
  object DataSource_Kvartira: TDataSource
    DataSet = FDQuery_Kvartira
    Left = 239
    Top = 302
  end
  object FDQuery_Stage: TFDQuery
    Connection = FDConnection_osn
    SQL.Strings = (
      'SELECT *  FROM St_tbl')
    Left = 140
    Top = 361
    object FDQuery_Stageid: TFDAutoIncField
      FieldName = 'id'
      Origin = 'id'
      ProviderFlags = [pfInWhere, pfInKey]
    end
    object FDQuery_StageFrom_Date: TDateTimeField
      FieldName = 'From_Date'
      Origin = 'From_Date'
    end
    object FDQuery_StageTo_Date: TDateTimeField
      FieldName = 'To_Date'
      Origin = 'To_Date'
    end
    object FDQuery_StageCompany: TWideMemoField
      FieldName = 'Company'
      Origin = 'Company'
      OnGetText = FDQuery_StageCompanyGetText
      BlobType = ftWideMemo
    end
    object FDQuery_StageDney_St: TIntegerField
      FieldName = 'Dney_St'
      Origin = 'Dney_St'
    end
    object FDQuery_StageGod_St: TStringField
      FieldName = 'God_St'
      Origin = 'God_St'
      Size = 100
    end
  end
  object DataSource_Stage: TDataSource
    DataSet = FDQuery_Stage
    OnDataChange = DataSource_StageDataChange
    Left = 236
    Top = 364
  end
  object FDQuery_Log: TFDQuery
    Connection = FDConnection_osn
    SQL.Strings = (
      'SELECT * FROM Logins')
    Left = 132
    Top = 425
    object FDQuery_Logid: TFDAutoIncField
      DisplayWidth = 2
      FieldName = 'id'
      Origin = 'id'
      ProviderFlags = [pfInWhere, pfInKey]
    end
    object FDQuery_Logurl: TWideMemoField
      DisplayWidth = 10
      FieldName = 'url'
      Origin = 'url'
      OnGetText = FDQuery_LogurlGetText
      BlobType = ftWideMemo
    end
    object FDQuery_Logusername: TWideMemoField
      DisplayWidth = 10
      FieldName = 'username'
      Origin = 'username'
      OnGetText = FDQuery_LogusernameGetText
      BlobType = ftWideMemo
    end
    object FDQuery_Logpassword: TStringField
      DisplayWidth = 1024
      FieldName = 'password'
      Origin = 'password'
      Size = 1024
    end
    object FDQuery_LoghttpRealm: TStringField
      DisplayWidth = 50
      FieldName = 'httpRealm'
      Origin = 'httpRealm'
      Size = 50
    end
    object FDQuery_LogformActionOrigin: TStringField
      DisplayWidth = 50
      FieldName = 'formActionOrigin'
      Origin = 'formActionOrigin'
      Size = 50
    end
    object FDQuery_Logguid: TStringField
      DisplayWidth = 50
      FieldName = 'guid'
      Origin = 'guid'
      Size = 50
    end
    object FDQuery_LogtimeCreated: TIntegerField
      DisplayWidth = 10
      FieldName = 'timeCreated'
      Origin = 'timeCreated'
    end
    object FDQuery_LogtimeLastUsed: TIntegerField
      DisplayWidth = 10
      FieldName = 'timeLastUsed'
      Origin = 'timeLastUsed'
    end
    object FDQuery_LogtimePasswordChanged: TIntegerField
      DisplayWidth = 17
      FieldName = 'timePasswordChanged'
      Origin = 'timePasswordChanged'
    end
  end
  object DataSource_Log: TDataSource
    DataSet = FDQuery_Log
    Left = 236
    Top = 425
  end
  object FDQuery_ins: TFDQuery
    Connection = FDConnection_osn
    Left = 385
    Top = 182
  end
  object DataSource_Edit: TDataSource
    DataSet = FDQuery_Edit
    OnDataChange = DataSource_EditDataChange
    Left = 232
    Top = 496
  end
  object FDQuery_Edit: TFDQuery
    Connection = FDConnection_osn
    SQL.Strings = (
      'select * from Bill_New')
    Left = 136
    Top = 496
  end
  object FDQuery_Voda: TFDQuery
    Connection = FDConnection_osn
    SQL.Strings = (
      'select * from Voda')
    Left = 136
    Top = 568
  end
  object DataSource_Voda: TDataSource
    DataSet = FDQuery_Voda
    Left = 240
    Top = 568
  end
  object FDQuery_C3: TFDQuery
    Connection = FDConnection_osn
    Left = 264
    Top = 24
  end
  object FDQuery_EPG: TFDQuery
    Connection = FDConnection_osn
    SQL.Strings = (
      'select * from EGP_program')
    Left = 392
    Top = 40
  end
  object FDQuery_EPG_channel: TFDQuery
    Connection = FDConnection_osn
    Left = 384
    Top = 112
  end
  object FDQuery_EPG_CUR: TFDQuery
    IndexesActive = False
    Connection = FDConnection_osn
    SQL.Strings = (
      'Select * From EPG_CUR')
    Left = 384
    Top = 240
    object FDQuery_EPG_CURid: TFDAutoIncField
      FieldName = 'id'
      Origin = 'id'
      ProviderFlags = [pfInWhere, pfInKey]
    end
    object FDQuery_EPG_CURchannel_en: TWideMemoField
      FieldName = 'channel_en'
      Origin = 'channel_en'
      OnGetText = FDQuery_EPG_CURchannel_enGetText
      BlobType = ftWideMemo
    end
    object FDQuery_EPG_CURchannel_ru: TWideMemoField
      FieldName = 'channel_ru'
      Origin = 'channel_ru'
      OnGetText = FDQuery_EPG_CURchannel_ruGetText
      BlobType = ftWideMemo
    end
    object FDQuery_EPG_CURperedacha: TWideMemoField
      FieldName = 'peredacha'
      Origin = 'peredacha'
      OnGetText = FDQuery_EPG_CURperedachaGetText
      BlobType = ftWideMemo
    end
    object FDQuery_EPG_CURT_start: TDateTimeField
      FieldName = 'T_start'
      Origin = 'T_start'
    end
    object FDQuery_EPG_CURT_stop: TDateTimeField
      FieldName = 'T_stop'
      Origin = 'T_stop'
    end
    object FDQuery_EPG_CURT_start_time: TWideMemoField
      FieldName = 'T_start_time'
      Origin = 'T_start_time'
      OnGetText = FDQuery_EPG_CURT_start_timeGetText
      BlobType = ftWideMemo
    end
    object FDQuery_EPG_CURT_stop_time: TWideMemoField
      FieldName = 'T_stop_time'
      Origin = 'T_stop_time'
      OnGetText = FDQuery_EPG_CURT_stop_timeGetText
      BlobType = ftWideMemo
    end
    object FDQuery_EPG_CURT_START_STOP: TIntegerField
      FieldName = 'T_START_STOP'
      OnGetText = FDQuery_EPG_CURT_START_STOPGetText
    end
  end
  object DataSource_EPG_CUR: TDataSource
    DataSet = FDQuery_EPG_CUR
    Left = 504
    Top = 232
  end
  object FDQuery_Den_Inp: TFDQuery
    Connection = FDConnection_osn
    Left = 360
    Top = 456
  end
  object FDQuery_Dengi_Pens: TFDQuery
    Connection = FDConnection_osn
    Left = 376
    Top = 392
  end
  object FDQuery_Dengi_Zarpl: TFDQuery
    Connection = FDConnection_osn
    SQL.Strings = (
      'select * from Vipiska_Zarpl')
    Left = 368
    Top = 312
    object FDQuery_Dengi_Zarplid: TFDAutoIncField
      FieldName = 'id'
      Origin = 'id'
      ProviderFlags = [pfInWhere, pfInKey]
    end
    object FDQuery_Dengi_ZarplSumma: TFloatField
      FieldName = 'Summa'
      Origin = 'Summa'
    end
    object FDQuery_Dengi_ZarplNomer: TIntegerField
      FieldName = 'Nomer'
      Origin = 'Nomer'
    end
    object FDQuery_Dengi_ZarplData_Op: TDateTimeField
      FieldName = 'Data_Op'
      Origin = 'Data_Op'
    end
    object FDQuery_Dengi_ZarplDetali: TWideMemoField
      FieldName = 'Detali'
      Origin = 'Detali'
      OnGetText = FDQuery_Dengi_ZarplDetaliGetText
      BlobType = ftWideMemo
    end
    object FDQuery_Dengi_ZarplOstatok_Zarpl: TFloatField
      FieldName = 'Ostatok_Zarpl'
      Origin = 'Ostatok_Zarpl'
    end
  end
  object DataSource_Dengi_Zarpl: TDataSource
    DataSet = FDQuery_Dengi_Zarpl
    OnDataChange = DataSource_Dengi_ZarplDataChange
    Left = 504
    Top = 312
  end
end
