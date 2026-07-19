unit Bill_Opl_Sheta;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs,
  // чтобы не было хинтов
  FireDAC.Stan.Param,
  Vcl.StdCtrls, Vcl.Samples.Spin,
  Vcl.ComCtrls;

type
  TForm7 = class(TForm)
    Button1: TButton;
    Button2: TButton;
    Label2: TLabel;
    Edit1: TEdit;
    Label1: TLabel;
    SpinEdit1: TSpinEdit;
    DateTimePicker1: TDateTimePicker;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Edit2: TEdit;
    Edit3: TEdit;
    Button3: TButton;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form7: TForm7;

implementation

{$R *.dfm}

uses New_Bill_Tab, DM;



procedure TForm7.Button1Click(Sender: TObject);
begin
Form7.Close;
end;

procedure TForm7.Button2Click(Sender: TObject);
var
Text_Sql     : string;
sel_row   : integer;
//Date_Oplati  : string ;
//Summa_Oplati : integer;
//Dolg         : integer;
//Dolg_pred    : integer;
//Za_chto      : integer;
//Status       : string;
seg_data     : string;
begin
//
sel_row  := DataModule1.DataSource_Kvartira.DataSet.RecNo;

Text_SQL := 'UPDATE Kvartira SET ';
Text_SQL :=Text_SQL + 'Date_Oplati = :p1,';
Text_SQL :=Text_SQL + ' Summa_Oplati  = :p2,';
Text_SQL :=Text_SQL + ' Za_chto = :p3,';
Text_SQL :=Text_SQL + ' Status = :p4,';
Text_SQL :=Text_SQL + ' Dolg = :p5';
Text_SQL :=Text_SQL + '  WHERE id = :p12;';
//Text_SQL :=Text_SQL + 'Date_Oplati ='

seg_data :=FormatDateTime('yyyy-mm-dd',StrToDate(Form1.today_str) );
//Dolg :=

DataModule1.FDQuery_Kvartira.Active :=false;
DataModule1.FDQuery_Kvartira.SQL.Clear;
DataModule1.FDQuery_Kvartira.SQL.Add(Text_SQL);


DataModule1.FDQuery_Kvartira.Params.ParamByName('p1').Value :=seg_data;
DataModule1.FDQuery_Kvartira.Params.ParamByName('p2').Value :=Form7.SpinEdit1.Value;
DataModule1.FDQuery_Kvartira.Params.ParamByName('p3').Value :=Form7.Edit1.Text;
DataModule1.FDQuery_Kvartira.Params.ParamByName('p4').Value :=  'Оплачен';
DataModule1.FDQuery_Kvartira.Params.ParamByName('p5').Value :=  Form7.Edit2.Text;
DataModule1.FDQuery_Kvartira.Params.ParamByName('p12').Value := sel_row ;
// нельзя чтобы FDQuery_Kvartira была активна статически
DataModule1.FDQuery_Kvartira.ExecSQL;    //иначе тут будет ошибка Data Base is locked

Text_SQL := 'SELECT * FROM Kvartira';
DataModule1.FDQuery_Kvartira.Active :=false;
DataModule1.FDQuery_Kvartira.SQL.Clear;
DataModule1.FDQuery_Kvartira.SQL.Add(Text_SQL);
DataModule1.FDQuery_Kvartira.Active :=true;

Form7.Close;
end;

procedure TForm7.Button3Click(Sender: TObject);
var
Text_Sql     : string;
sel_row   : integer;
//Date_Oplati  : string ;
//Summa_Oplati : integer;
//Dolg         : integer;
//Dolg_pred    : integer;
//Za_chto      : integer;
//Status       : string;
seg_data     : string;
begin
//
sel_row  := DataModule1.DataSource_Kvartira.DataSet.RecNo;

Text_SQL := 'UPDATE Kvartira SET ';
Text_SQL :=Text_SQL + 'Date_Oplati = :p1,';
Text_SQL :=Text_SQL + ' Summa_Oplati  = :p2,';
Text_SQL :=Text_SQL + ' Za_chto = :p3,';
Text_SQL :=Text_SQL + ' Status = :p4,';
Text_SQL :=Text_SQL + ' Dolg = :p5';
Text_SQL :=Text_SQL + '  WHERE id = :p12;';
//Text_SQL :=Text_SQL + 'Date_Oplati ='

seg_data :=FormatDateTime('yyyy-mm-dd',StrToDate(Form1.today_str) );
//Dolg :=

DataModule1.FDQuery_Kvartira.Active :=false;
DataModule1.FDQuery_Kvartira.SQL.Clear;
DataModule1.FDQuery_Kvartira.SQL.Add(Text_SQL);


DataModule1.FDQuery_Kvartira.Params.ParamByName('p1').Value :=seg_data;
DataModule1.FDQuery_Kvartira.Params.ParamByName('p2').Value :=Form7.SpinEdit1.Value;
DataModule1.FDQuery_Kvartira.Params.ParamByName('p3').Value :=Form7.Edit1.Text;
DataModule1.FDQuery_Kvartira.Params.ParamByName('p4').Value := '';
DataModule1.FDQuery_Kvartira.Params.ParamByName('p5').Value :=  Form7.Edit2.Text;
DataModule1.FDQuery_Kvartira.Params.ParamByName('p12').Value := sel_row ;
// нельзя чтобы FDQuery_Kvartira была активна статически
DataModule1.FDQuery_Kvartira.ExecSQL;    //иначе тут будет ошибка Data Base is locked

Text_SQL := 'SELECT * FROM Kvartira';
DataModule1.FDQuery_Kvartira.Active :=false;
DataModule1.FDQuery_Kvartira.SQL.Clear;
DataModule1.FDQuery_Kvartira.SQL.Add(Text_SQL);
DataModule1.FDQuery_Kvartira.Active :=true;

Form7.Close;

end;

procedure TForm7.FormClose(Sender: TObject; var Action: TCloseAction);
begin
Form1.Show;
end;

end.
