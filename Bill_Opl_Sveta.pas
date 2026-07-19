unit Bill_Opl_Sveta;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs,
  // чтобы не было хинтов
  FireDAC.Stan.Param,
  //
   Vcl.StdCtrls;

type
  TForm8 = class(TForm)
    Button1: TButton;
    Edit1: TEdit;
    Label11: TLabel;
    Edit2: TEdit;
    Label1: TLabel;
    Edit3: TEdit;
    Label2: TLabel;
    Edit4: TEdit;
    Label3: TLabel;
    Button2: TButton;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form8: TForm8;

implementation

{$R *.dfm}

uses New_Bill_Tab, DM;



procedure TForm8.Button1Click(Sender: TObject);
var
Text_SQL : string;
begin
   Text_SQL := 'Select * From Svet';
   DataModule1.FDQuery_Svet.Active := false;
   DataModule1.FDQuery_Svet.SQL.Clear;
   DataModule1.FDQuery_Svet.SQL.Add(Text_SQL);
   DataModule1.FDQuery_Svet.Active := true;
   DataModule1.FDQuery_Svet.Last;

Form8.Close;
end;

procedure TForm8.Button2Click(Sender: TObject);
var
Text_Sql : string;
begin
   Text_SQL := 'update Svet set Date_Sn_Pokaz =:p1,  ';
   Text_SQL :=Text_SQL + 'Pokaz = :p2,  ';
   Text_SQL :=Text_SQL + 'Date_Opl= :p3,  ';
   Text_SQL :=Text_SQL + 'Razn= :p4  ';
   Text_SQL :=Text_SQL +'where id = '+ IntToStr(Form1.sel_row);

   DataModule1.FDQuery_Svet.Active := false;
   DataModule1.FDQuery_Svet.SQL.Clear;
   DataModule1.FDQuery_Svet.SQL.Add(Text_SQL);

   DataModule1.FDQuery_Svet.ParamByName('p1').AsString :=
   FormatDateTime('yyyy-mm-dd',StrToDate(Form8.Edit1.Text));

   DataModule1.FDQuery_Svet.ParamByName('p2').AsString :=
   Form8.Edit2.Text;

   DataModule1.FDQuery_Svet.ParamByName('p3').AsString :=
   FormatDateTime('yyyy-mm-dd',StrToDate(Form8.Edit3.Text));

   DataModule1.FDQuery_Svet.ParamByName('p4').AsString :=
   Form8.Edit4.Text;

   DataModule1.FDQuery_Svet.ExecSQL;

   Text_SQL := 'Select * From Svet';
   DataModule1.FDQuery_Svet.Active := false;
   DataModule1.FDQuery_Svet.SQL.Clear;
   DataModule1.FDQuery_Svet.SQL.Add(Text_SQL);
   DataModule1.FDQuery_Svet.Active := true;
   DataModule1.FDQuery_Svet.Last;

  Form8.Close;
end;

procedure TForm8.FormClose(Sender: TObject; var Action: TCloseAction);
begin
Form1.Show;
end;

procedure TForm8.FormShow(Sender: TObject);
var
 Text_SQL      : string;
 Date_Sn_Pokaz : string;
 Pokaz         : integer;
 Pokaz_pr      : integer;
 Date_Opl      : string;
 Razn          : string;
begin
   Text_SQL := 'Select * From Svet';
   Text_SQL := Text_SQL + '  where id = ' +IntToStr(Form1.sel_row); ;
   DataModule1.FDQuery_Svet.Active := false;
   DataModule1.FDQuery_Svet.SQL.Clear;
   DataModule1.FDQuery_Svet.SQL.Add(Text_SQL);
   DataModule1.FDQuery_Svet.Active := true;


 Date_Sn_Pokaz := DataModule1.FDQuery_Svet.FieldByName('Date_Sn_Pokaz').AsString;
 Pokaz := DataModule1.FDQuery_Svet.FieldByName('Pokaz').AsInteger;
 Date_Opl := DataModule1.FDQuery_Svet.FieldByName('Date_Opl').AsString;
//================================
   Text_SQL := 'Select * From Svet';
   Text_SQL := Text_SQL + '  where id = ' +IntToStr(Form1.sel_row-1); ;
   DataModule1.FDQuery_Svet.Active := false;
   DataModule1.FDQuery_Svet.SQL.Clear;
   DataModule1.FDQuery_Svet.SQL.Add(Text_SQL);
   DataModule1.FDQuery_Svet.Active := true;
Pokaz_pr  := DataModule1.FDQuery_Svet.FieldByName('Pokaz').AsInteger;
Razn := IntToStr(Pokaz - Pokaz_pr) ;

Form8.Edit1.Text := Date_Sn_Pokaz;
Form8.Edit2.Text := IntToStr(Pokaz);
if Date_Opl ='' then
  begin
  Form8.Edit3.Text := Form1.today_str;
  end
else
  begin
  Form8.Edit3.Text := Date_Opl;
  end;
Form8.Edit4.Text := Razn;

end;

end.
