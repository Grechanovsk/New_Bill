unit Bill_Shet;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms,
  // чтобы не было хинтов
  FireDAC.Stan.Param,

  Vcl.Dialogs, Vcl.StdCtrls, Vcl.Samples.Spin;

type
  TForm6 = class(TForm)
    Label1: TLabel;
    SpinEdit1: TSpinEdit;
    Label2: TLabel;
    Edit1: TEdit;
    Button2: TButton;
    Button1: TButton;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Button1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Button2Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form6: TForm6;

implementation

{$R *.dfm}

uses New_Bill_Tab, DM;



procedure TForm6.Button1Click(Sender: TObject);
begin
Form6.Close;
end;

procedure TForm6.Button2Click(Sender: TObject);
var
Text_SQL : string;
Za_chto  : string;
Summa_Oplati : integer;
begin
//
   Text_SQL := 'insert into Kvartira    ';
   Text_SQL :=Text_SQL +'(Summa_Oplati,Za_chto) ';
   Text_SQL :=Text_SQL + ' values (:p1,:p2)';

   DataModule1.FDQuery_Kvartira.Active := false;
   DataModule1.FDQuery_Kvartira.SQL.Clear;
   DataModule1.FDQuery_Kvartira.SQL.Add(Text_SQL);

   Summa_Oplati := Form6.SpinEdit1.Value;
   Za_chto := Form6.Edit1.Text;
   DataModule1.FDQuery_Kvartira.Params.ParamByName('p1').Value := Summa_Oplati;
   DataModule1.FDQuery_Kvartira.Params.ParamByName('p2').Value := Za_chto;
   DataModule1.FDQuery_Kvartira.ExecSQL;

   Text_SQL := 'select * from Kvartira    ';
   DataModule1.FDQuery_Kvartira.Active := false;
   DataModule1.FDQuery_Kvartira.SQL.Clear;
   DataModule1.FDQuery_Kvartira.SQL.Add(Text_SQL);
   DataModule1.FDQuery_Kvartira.Active := true;

//   Form1.FDQuery_Kvartira.Last;
   Form6.Close;

end;

procedure TForm6.FormClose(Sender: TObject; var Action: TCloseAction);
begin
Form1.Show;
end;

procedure TForm6.FormCreate(Sender: TObject);
begin
Form6.SpinEdit1.Value := 0 ;
Form6.Edit1.Text := '';
end;

end.
