unit DM;

interface

uses
  System.SysUtils, System.Classes, Vcl.ExtCtrls, FireDAC.Stan.Intf,
  FireDAC.Stan.Option, FireDAC.Stan.Error, FireDAC.UI.Intf, FireDAC.Phys.Intf,
  FireDAC.Stan.Def, FireDAC.Stan.Pool, FireDAC.Stan.Async, FireDAC.Phys,
  FireDAC.Phys.SQLite, FireDAC.Phys.SQLiteDef, FireDAC.Stan.ExprFuncs,
  FireDAC.Phys.SQLiteWrapper.Stat, FireDAC.VCLUI.Wait, Data.DB,
  FireDAC.Comp.Client, Vcl.Menus, FireDAC.Stan.Param, FireDAC.DatS,
  FireDAC.DApt.Intf, FireDAC.DApt,
  //использование внешних модулей
  DelphiNotesAge,
  // для использования временных утилит
  DateUtils,
  //
   FireDAC.Comp.DataSet;

type
  TDataModule1 = class(TDataModule)
    FDConnection_osn: TFDConnection;
    FDQuery_old_1: TFDQuery;
    FDQuery_osn: TFDQuery;
    FDQuery_Svet: TFDQuery;
    DataSource_osn: TDataSource;
    FDQuery_opl: TFDQuery;
    DataSource_opl: TDataSource;
    FDQuery_excel: TFDQuery;
    DataSource_excel: TDataSource;
    DataSource_Svet: TDataSource;
    FDQuery_Kvartira: TFDQuery;
    DataSource_Kvartira: TDataSource;
    FDQuery_Stage: TFDQuery;
    FDQuery_Stageid: TFDAutoIncField;
    FDQuery_StageFrom_Date: TDateTimeField;
    FDQuery_StageTo_Date: TDateTimeField;
    FDQuery_StageCompany: TWideMemoField;
    FDQuery_StageDney_St: TIntegerField;
    FDQuery_StageGod_St: TStringField;
    DataSource_Stage: TDataSource;
    FDQuery_Log: TFDQuery;
    FDQuery_Logid: TFDAutoIncField;
    FDQuery_Logurl: TWideMemoField;
    FDQuery_Logusername: TWideMemoField;
    FDQuery_Logpassword: TStringField;
    FDQuery_LoghttpRealm: TStringField;
    FDQuery_LogformActionOrigin: TStringField;
    FDQuery_Logguid: TStringField;
    FDQuery_LogtimeCreated: TIntegerField;
    FDQuery_LogtimeLastUsed: TIntegerField;
    FDQuery_LogtimePasswordChanged: TIntegerField;
    DataSource_Log: TDataSource;
    FDQuery_ins: TFDQuery;
    DataSource_Edit: TDataSource;
    FDQuery_Edit: TFDQuery;
    FDQuery_Voda: TFDQuery;
    DataSource_Voda: TDataSource;
    FDQuery_C3: TFDQuery;
    FDQuery_EPG: TFDQuery;
    FDQuery_EPG_channel: TFDQuery;
    FDQuery_EPG_CUR: TFDQuery;
    FDQuery_EPG_CURid: TFDAutoIncField;
    FDQuery_EPG_CURchannel_en: TWideMemoField;
    FDQuery_EPG_CURchannel_ru: TWideMemoField;
    FDQuery_EPG_CURperedacha: TWideMemoField;
    FDQuery_EPG_CURT_start: TDateTimeField;
    FDQuery_EPG_CURT_stop: TDateTimeField;
    FDQuery_EPG_CURT_start_time: TWideMemoField;
    FDQuery_EPG_CURT_stop_time: TWideMemoField;
    DataSource_EPG_CUR: TDataSource;
    FDQuery_EPG_CURT_START_STOP: TIntegerField;
    FDQuery_Den_Inp: TFDQuery;
    FDQuery_Dengi_Pens: TFDQuery;
    FDQuery_Dengi_Zarpl: TFDQuery;
    DataSource_Dengi_Zarpl: TDataSource;
    FDQuery_Dengi_Zarplid: TFDAutoIncField;
    FDQuery_Dengi_ZarplSumma: TFloatField;
    FDQuery_Dengi_ZarplNomer: TIntegerField;
    FDQuery_Dengi_ZarplData_Op: TDateTimeField;
    FDQuery_Dengi_ZarplDetali: TWideMemoField;
    FDQuery_Dengi_ZarplOstatok_Zarpl: TFloatField;
    procedure DataSource_StageDataChange(Sender: TObject; Field: TField);
    procedure FDQuery_StageCompanyGetText(Sender: TField; var Text: string;
      DisplayText: Boolean);
    procedure FDQuery_LogurlGetText(Sender: TField; var Text: string;
      DisplayText: Boolean);
    procedure FDQuery_LogusernameGetText(Sender: TField; var Text: string;
      DisplayText: Boolean);
    procedure DataSource_EditDataChange(Sender: TObject; Field: TField);
    procedure FDQuery_EPG_CURchannel_enGetText(Sender: TField; var Text: string;
      DisplayText: Boolean);
    procedure FDQuery_EPG_CURchannl_ruGetText(Sender: TField; var Text: string;
      DisplayText: Boolean);
    procedure FDQuery_EPG_CURperedachaGetText(Sender: TField; var Text: string;
      DisplayText: Boolean);
    procedure FDQuery_EPG_CURT_start_timeGetText(Sender: TField;
      var Text: string; DisplayText: Boolean);
    procedure FDQuery_EPG_CURT_stop_timeGetText(Sender: TField;
      var Text: string; DisplayText: Boolean);
    procedure FDQuery_EPG_CURTSTARTSTOPGetText(Sender: TField; var Text: string;
      DisplayText: Boolean);
    procedure FDQuery_EPG_CURchannel_ruGetText(Sender: TField; var Text: string;
      DisplayText: Boolean);
    procedure FDQuery_EPG_CURT_START_STOPGetText(Sender: TField;
      var Text: string; DisplayText: Boolean);
    procedure FDQuery_DengiVid_PlatGetText(Sender: TField; var Text: string;
      DisplayText: Boolean);
    procedure FDQuery_DengiType_PlatGetText(Sender: TField; var Text: string;
      DisplayText: Boolean);
    procedure FDQuery_Dengi_ZarplidGetText(Sender: TField; var Text: string;
      DisplayText: Boolean);
    procedure FDQuery_DengiChisloGetText(Sender: TField; var Text: string;
      DisplayText: Boolean);
    procedure FDQuery_Dengi_ZarplSummaGetText(Sender: TField; var Text: string;
      DisplayText: Boolean);
    procedure FDQuery_DengiOstatokGetText(Sender: TField; var Text: string;
      DisplayText: Boolean);
    procedure FDQuery_DengiKartaGetText(Sender: TField; var Text: string;
      DisplayText: Boolean);
    procedure FDQuery_Dengi_ZarplDetaliGetText(Sender: TField; var Text: string;
      DisplayText: Boolean);
   // procedure FDQuery_Dengi_ZarplContragentGetText(Sender: TField; var Text: string;
   //   DisplayText: Boolean);
    procedure DataSource_Dengi_ZarplDataChange(Sender: TObject; Field: TField);
    procedure FDQuery_Dengi_ZarplData_OpGetText(Sender: TField; var Text: string;
      DisplayText: Boolean);
    procedure FDQuery_Dengi_ZarplNomerGetText(Sender: TField; var Text: string;
      DisplayText: Boolean);
    procedure FDQuery_Dengi_PensDetaliGetText(Sender: TField; var Text: string;
      DisplayText: Boolean);
   // procedure FDQuery_Dengi_PensContragentGetText(Sender: TField;
   //   var Text: string; DisplayText: Boolean);
    procedure DataSource_Dengi_PensDataChange(Sender: TObject; Field: TField);
  private
    { Private declarations }
  public
    { Public declarations }
   sel_row_St : integer;
   Dney_st    : string;
   God_St  : string;

  end;

var
  DataModule1: TDataModule1;

implementation

{%CLASSGROUP 'Vcl.Controls.TControl'}

uses New_Bill_Tab, Bill3;

{$R *.dfm}

procedure TDataModule1.DataSource_Dengi_PensDataChange(Sender: TObject;
  Field: TField);
begin
 //
 Form1.Memo7.Text := DataModule1.FDQuery_Dengi_Pens.FieldByName('Detali').AsString;
 Form1.Memo6.Text := DataModule1.FDQuery_Dengi_Pens.FieldByName('Contragent').AsString;

end;

procedure TDataModule1.DataSource_Dengi_ZarplDataChange(Sender: TObject;
  Field: TField);
begin
 //
 Form1.Memo7.Text := DataModule1.FDQuery_Dengi_Zarpl.FieldByName('Detali').AsString;
// Form1.Memo6.Text := DataModule1.FDQuery_Dengi_Zarpl.FieldByName('Contragent').AsString;
end;

procedure TDataModule1.DataSource_EditDataChange(Sender: TObject;
  Field: TField);
begin

Form3.Edit1.Text :=
  DataModule1.FDQuery_Edit.FieldByName('Data').AsString;
Form3.Edit7.Text :=
  DataModule1.FDQuery_Edit.FieldByName('Ostatok').AsString;
Form3.Edit2.Text :=
  DataModule1.FDQuery_Edit.FieldByName('Travm').AsString;
Form3.Edit3.Text :=
  DataModule1.FDQuery_Edit.FieldByName('Gips').AsString;
Form3.Edit4.Text :=
  DataModule1.FDQuery_Edit.FieldByName('Zapl_inet').AsString;
Form3.Edit5.Text :=
  DataModule1.FDQuery_Edit.FieldByName('Sn_Gips').AsString;
Form3.Edit6.Text :=
  DataModule1.FDQuery_Edit.FieldByName('Den_Gizni').AsString;

Form3.ComboBox1.Text :=
  DataModule1.FDQuery_Edit.FieldByName('Pil').AsString;
Form3.ComboBox2.Text :=
  DataModule1.FDQuery_Edit.FieldByName('Sobitie').AsString;
Form3.ComboBox3.Text :=
  DataModule1.FDQuery_Edit.FieldByName('DenNed').AsString;


end;

procedure TDataModule1.DataSource_StageDataChange(Sender: TObject;
  Field: TField);
var
  Stage      : integer;
  From_Date  : Tdate;
  To_Date    : Tdate;
 begin
  Form1.DateTimePicker3.Date :=DataModule1.FDQuery_Stage.FieldByName('From_date').AsDateTime;
  Form1.DateTimePicker4.Date :=DataModule1.FDQuery_Stage.FieldByName('To_date').AsDateTime;
  Form1.Memo3.Text := DataModule1.FDQuery_Stage.FieldByName('Company').AsString ;

  From_Date  := DataModule1.FDQuery_Stage.FieldByName('From_Date').AsDateTime;
  To_Date    := DataModule1.FDQuery_Stage.FieldByName('To_Date').AsDateTime;
  Stage := DaysBetween(To_Date, From_Date  ) ;
  Form1.Edit11.Text := IntToStr(Stage);
  Form1.Edit13.Text := FormatAge(Now - StrToInt(Form1.Edit11.Text), Now, True);// 27 л. 7 мес. 24 д.
//запоминаем значения переменных
  sel_row_St := DataModule1.DataSource_Stage.DataSet.RecNo;
  Dney_st := Form1.Edit11.Text;
  God_St  := Form1.Edit13.Text;


end;

procedure TDataModule1.FDQuery_DengiChisloGetText(Sender: TField;
  var Text: string; DisplayText: Boolean);
begin
Text := Sender.AsString;
end;

{procedure TDataModule1.FDQuery_Dengi_PensContragentGetText(Sender: TField;
  var Text: string; DisplayText: Boolean);
begin
Text := Sender.AsString;
end;}

procedure TDataModule1.FDQuery_Dengi_PensDetaliGetText(Sender: TField;
  var Text: string; DisplayText: Boolean);
begin
  Text := Sender.AsString;
end;
{
procedure TDataModule1.FDQuery_Dengi_ZarplContragentGetText(Sender: TField;
  var Text: string; DisplayText: Boolean);
begin
Text := Sender.AsString;
end;
 }
procedure TDataModule1.FDQuery_Dengi_ZarplData_OpGetText(Sender: TField;
  var Text: string; DisplayText: Boolean);
begin
Text := Sender.AsString;
end;

procedure TDataModule1.FDQuery_Dengi_ZarplDetaliGetText(Sender: TField;
  var Text: string; DisplayText: Boolean);
begin
Text := Sender.AsString;
end;

procedure TDataModule1.FDQuery_Dengi_ZarplidGetText(Sender: TField; var Text: string;
  DisplayText: Boolean);
begin
Text := Sender.AsString;
end;

procedure TDataModule1.FDQuery_Dengi_ZarplNomerGetText(Sender: TField;
  var Text: string; DisplayText: Boolean);
begin
  Text := Sender.AsString;
end;

procedure TDataModule1.FDQuery_DengiKartaGetText(Sender: TField;
  var Text: string; DisplayText: Boolean);
begin
Text := Sender.AsString;
end;

procedure TDataModule1.FDQuery_DengiOstatokGetText(Sender: TField;
  var Text: string; DisplayText: Boolean);
begin
Text := Sender.AsString;
end;

procedure TDataModule1.FDQuery_Dengi_ZarplSummaGetText(Sender: TField;
  var Text: string; DisplayText: Boolean);
begin
Text := Sender.AsString ;
end;

procedure TDataModule1.FDQuery_DengiType_PlatGetText(Sender: TField;
  var Text: string; DisplayText: Boolean);
begin
  Text := Sender.AsString;
end;

procedure TDataModule1.FDQuery_DengiVid_PlatGetText(Sender: TField;
  var Text: string; DisplayText: Boolean);
begin
  Text := Sender.AsString;
end;

procedure TDataModule1.FDQuery_EPG_CURchannel_enGetText(Sender: TField;
  var Text: string; DisplayText: Boolean);
begin
 Text := Sender.AsString;
end;

procedure TDataModule1.FDQuery_EPG_CURchannel_ruGetText(Sender: TField;
  var Text: string; DisplayText: Boolean);
begin
 Text := Sender.AsString;
end;

procedure TDataModule1.FDQuery_EPG_CURchannl_ruGetText(Sender: TField;
  var Text: string; DisplayText: Boolean);
begin
Text := Sender.AsString;
end;

procedure TDataModule1.FDQuery_EPG_CURperedachaGetText(Sender: TField;
  var Text: string; DisplayText: Boolean);
begin
Text := Sender.AsString;
end;

procedure TDataModule1.FDQuery_EPG_CURTSTARTSTOPGetText(Sender: TField;
  var Text: string; DisplayText: Boolean);
begin
 Text := Sender.AsString;
end;

procedure TDataModule1.FDQuery_EPG_CURT_START_STOPGetText(Sender: TField;
  var Text: string; DisplayText: Boolean);
begin
 Text := Sender.AsString;
end;

procedure TDataModule1.FDQuery_EPG_CURT_start_timeGetText(Sender: TField;
  var Text: string; DisplayText: Boolean);
begin
Text := Sender.AsString;
end;

procedure TDataModule1.FDQuery_EPG_CURT_stop_timeGetText(Sender: TField;
  var Text: string; DisplayText: Boolean);
begin
Text := Sender.AsString;
end;

procedure TDataModule1.FDQuery_LogurlGetText(Sender: TField; var Text: string;
  DisplayText: Boolean);
begin
 Text := Sender.AsString;
end;

procedure TDataModule1.FDQuery_LogusernameGetText(Sender: TField;
  var Text: string; DisplayText: Boolean);
begin
Text := Sender.AsString;
end;

procedure TDataModule1.FDQuery_StageCompanyGetText(Sender: TField;
  var Text: string; DisplayText: Boolean);
begin
   Text := Sender.AsString;
end;

end.
