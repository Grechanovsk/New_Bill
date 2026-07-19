unit Bill2;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ComCtrls, Vcl.StdCtrls,
  Vcl.Samples.Spin, FireDAC.Stan.Intf, FireDAC.Stan.Option, FireDAC.Stan.Param,
  FireDAC.Stan.Error, FireDAC.DatS, FireDAC.Phys.Intf, FireDAC.DApt.Intf,
  FireDAC.Stan.Async, FireDAC.DApt, Data.DB, FireDAC.Comp.DataSet,
  FireDAC.Comp.Client;

type
  TForm2 = class(TForm)
    DateTimePicker1: TDateTimePicker;
    Label1: TLabel;
    SpinEdit1: TSpinEdit;
    Label2: TLabel;
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form2: TForm2;

implementation

{$R *.dfm}

uses New_Bill_Tab, DM;



procedure TForm2.Button1Click(Sender: TObject);
var
Data_opl : string;
Text_SQL    : string;
//god_v   : string;
//mes_v  : string;
//day_v  : string;
//data_v : string;


begin
   Text_SQL := 'insert into Opl_table  ';
   Text_SQL :=Text_SQL +'(Data_opl,Sum_opl) ';
   Text_SQL :=Text_SQL + ' values (:p1,:p2)';

//   data_v := FormatDateTime('yyyy-mm-dd',c);

   DataModule1.FDQuery_ins.Active := false;
   DataModule1.FDQuery_ins.SQL.Clear;
   DataModule1.FDQuery_ins.SQL.Add(Text_SQL);
   Data_opl := FormatDateTime('yyyy-mm-dd',DateTimePicker1.Date);


   DataModule1.FDQuery_ins.Params.ParamByName('p1').Value := Data_opl;
   DataModule1.FDQuery_ins.Params.ParamByName('p2').Value := Form2.SpinEdit1.Value;
//DateTimeToJulianDate

   DataModule1.FDQuery_opl.Active := false;

   DataModule1.FDQuery_ins.ExecSQL;

   DataModule1.FDQuery_opl.Active := true;
   DataModule1.FDQuery_opl.Last;
   Form2.Close;

end;

procedure TForm2.Button2Click(Sender: TObject);
begin
Form2.SpinEdit1.Value := Form2.SpinEdit1.Value +10;
end;

procedure TForm2.Button3Click(Sender: TObject);
begin
Form2.SpinEdit1.Value := Form2.SpinEdit1.Value +100;
end;

procedure TForm2.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Form1.Show;
end;

procedure TForm2.FormCreate(Sender: TObject);
begin
//
//Form1.today_str
Form2.DateTimePicker1.Date := StrToDate(Form1.today_str) ;
end;

end.
