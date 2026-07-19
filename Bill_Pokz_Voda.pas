unit Bill_Pokz_Voda;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms,
  // чтобы не было хинтов
  FireDAC.Stan.Param,

  Vcl.Dialogs, Vcl.StdCtrls;

type
  TForm4 = class(TForm)
    Button2: TButton;
    Button1: TButton;
    Edit3: TEdit;
    Label2: TLabel;
    Edit2: TEdit;
    Label1: TLabel;
    Edit1: TEdit;
    Label11: TLabel;
    procedure Button1Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure Button2Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form4: TForm4;

implementation

{$R *.dfm}

uses New_Bill_Tab, DM;

procedure TForm4.Button1Click(Sender: TObject);
begin
Form4.Close;
end;

procedure TForm4.Button2Click(Sender: TObject);
var
Text_Sql : string;
tec_data : string;
begin
//
//Insert into Voda (Data,P_Gor,P_Hol) values ('2022-08-27',531,351)
Text_Sql := 'Insert into Voda (Data,P_Gor,P_Hol)';
Text_Sql := Text_Sql + ' values(:p1,:p2,:p3)';
tec_data  := FormatDateTime('yyyy-mm-dd', StrToDate(Form4.Edit1.Text));

DataModule1.FDQuery_Voda.Active := false;
DataModule1.FDQuery_Voda.SQL.Clear;
DataModule1.FDQuery_Voda.SQL.Add(Text_Sql);

DataModule1.FDQuery_Voda.Params.ParamByName('p1').Value :=tec_data;
DataModule1.FDQuery_Voda.Params.ParamByName('p2').Value :=Form4.Edit3.Text;
DataModule1.FDQuery_Voda.Params.ParamByName('p3').Value :=Form4.Edit2.Text;
DataModule1.FDQuery_Voda.ExecSQL;

Text_Sql := 'select * from Voda';
DataModule1.FDQuery_Voda.Active := false;
DataModule1.FDQuery_Voda.SQL.Clear;
DataModule1.FDQuery_Voda.SQL.Add(Text_Sql);
DataModule1.FDQuery_Voda.Active := true;
DataModule1.FDQuery_Voda.Last;


Form4.Close;

end;

procedure TForm4.FormClose(Sender: TObject; var Action: TCloseAction);
begin
Form1.Show;
end;

procedure TForm4.FormCreate(Sender: TObject);
begin

Form4.Edit1.Text := Form1.today_str;
Form4.Edit2.Clear;
Form4.Edit3.Clear;
Form4.Edit2.Color :=RGB(102, 204, 255);
Form4.Edit3.Color :=RGB(255, 51, 204);

end;

procedure TForm4.FormShow(Sender: TObject);
begin
Form4.Edit2.SetFocus;
end;

end.
