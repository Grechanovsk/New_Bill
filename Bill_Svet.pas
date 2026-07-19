unit Bill_Svet;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ComCtrls,
  FireDAC.Stan.Intf, FireDAC.Stan.Option, FireDAC.Stan.Param,
  FireDAC.Stan.Error,
  FireDAC.DatS, FireDAC.Phys.Intf, FireDAC.DApt.Intf,
  FireDAC.Stan.Async, FireDAC.DApt, Data.DB, FireDAC.Comp.DataSet,
  FireDAC.Comp.Client, FireDAC.UI.Intf, FireDAC.Stan.Def, FireDAC.Stan.Pool,
  FireDAC.Phys, FireDAC.Phys.SQLite, FireDAC.Phys.SQLiteDef,
  FireDAC.Stan.ExprFuncs, FireDAC.Phys.SQLiteWrapper.Stat, FireDAC.VCLUI.Wait;

type
  TForm5 = class(TForm)
    Button1: TButton;
    Edit1: TEdit;
    Label1: TLabel;
    Label2: TLabel;
    Button2: TButton;
    DateTimePicker1: TDateTimePicker;
    Label3: TLabel;
    Edit2: TEdit;
    procedure Button1Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure Edit1KeyPress(Sender: TObject; var Key: Char);
    procedure Button2Click(Sender: TObject);
    procedure Edit1Change(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    Pokaz         : integer;
    Pokaz_pr      : integer;

  end;

var
  Form5: TForm5;

implementation

{$R *.dfm}

uses New_Bill_Tab, DM;




procedure TForm5.Button2Click(Sender: TObject);
var
Text_SQL : string;
Date_Sn_Pokaz : string;
//Pokaz  : integer;
begin
//
   Text_SQL := 'insert into Svet  ';
   Text_SQL :=Text_SQL +'(Date_Sn_Pokaz,Pokaz) ';
   Text_SQL :=Text_SQL + ' values (:p1,:p2)';

//Date_Sn_Pokaz,Pokaz

   DataModule1.FDQuery_Svet.Active := false;
   DataModule1.FDQuery_Svet.SQL.Clear;
   DataModule1.FDQuery_Svet.SQL.Add(Text_SQL);
   Date_Sn_Pokaz := FormatDateTime('yyyy-mm-dd',Form5.DateTimePicker1.Date);
   DataModule1.FDQuery_Svet.Params.ParamByName('p1').Value := Date_Sn_Pokaz;
   DataModule1.FDQuery_Svet.Params.ParamByName('p2').Value := Form5.Edit1.Text;

//   Form1.FDQuery_Svet.Active := true;
   DataModule1.FDQuery_Svet.ExecSQL;

   Text_SQL := 'Select * From Svet';
   DataModule1.FDQuery_Svet.Active := false;
   DataModule1.FDQuery_Svet.SQL.Clear;
   DataModule1.FDQuery_Svet.SQL.Add(Text_SQL);
   DataModule1.FDQuery_Svet.Active := true;
   DataModule1.FDQuery_Svet.Last;

   Form5.Close;


end;

procedure TForm5.FormClose(Sender: TObject; var Action: TCloseAction);
begin
Form1.Show;
end;

procedure TForm5.Button1Click(Sender: TObject);
begin
Form5.Close;
Form1.Show;
end;

procedure TForm5.FormCreate(Sender: TObject);
begin
//
//Form5.Edit2.Text := Form1.today_str;
Form5.DateTimePicker1.Date := StrToDate(Form1.today_str);
//Form5.Edit1.SetFocus;
end;

procedure TForm5.FormActivate(Sender: TObject);
var
 Text_SQL     : string;
 //Pokaz         : integer;
 //Pokaz_pr      : integer;

begin
   Text_SQL := ' SELECT * FROM Svet  WHERE id=';
   Text_SQL := Text_SQL + '  (SELECT max(id) FROM Svet )' ; ;
   DataModule1.FDQuery_Svet.Active := false;
   DataModule1.FDQuery_Svet.SQL.Clear;
   DataModule1.FDQuery_Svet.SQL.Add(Text_SQL);
   DataModule1.FDQuery_Svet.Active := true;

    Pokaz_pr := DataModule1.FDQuery_Svet.FieldByName('Pokaz').AsInteger;
    Form5.Edit2.Text := IntToStr(Pokaz_pr);
Form5.Edit1.SetFocus;
end;

procedure TForm5.Edit1Change(Sender: TObject);
var
 Pokaz_tt : integer;
begin
//
Pokaz_tt := StrToInt(Form5.Edit1.Text) - Pokaz_pr  ;
Form5.Edit2.Text := IntToStr(Pokaz_tt);
end;

procedure TForm5.Edit1KeyPress(Sender: TObject; var Key: Char);
begin
if not (CharInSet(key,['0'..'9', #13])) then  Key := #0;
end;

end.
