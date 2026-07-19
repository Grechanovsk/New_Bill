unit Edit_Zam;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, FireDAC.Stan.Intf, FireDAC.Stan.Option,
  FireDAC.Stan.Param, FireDAC.Stan.Error, FireDAC.DatS, FireDAC.Phys.Intf,
  FireDAC.DApt.Intf, FireDAC.Stan.Async, FireDAC.DApt, Data.DB,
  FireDAC.Comp.DataSet,
  //

  FireDAC.Comp.Client, Vcl.StdCtrls;

type
  TForm9 = class(TForm)
    Button1: TButton;
    Memo1: TMemo;
    Button2: TButton;
    FDQuery1: TFDQuery;
    Label48: TLabel;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure Label48Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    kkk: integer;
    Text_sql : string;
    zamet : string;
    zametki : TStringList;
    s      : string;
    //from Form1
    Layout: array[0.. KL_NAMELENGTH] of char;

  end;
const
  Trawm_ish = '09.05.2021';
  Sn_Gips_ish = '29.07.2021';
  Gips_ish = '13.05.2021';
  //keyboard
  CNT_LAYOUT = 2; // количество известных раскладок
  ENGLISH = $409;
  RUSSIAN = $419;

  TKbdValue : array [1..CNT_LAYOUT] of LongWord =
                ( ENGLISH,
                  RUSSIAN
                );
  TKbdDisplayNames : array [1..CNT_LAYOUT] of string =
                ('English',
                 'Русский'
                );


var
  Form9: TForm9;

implementation

{$R *.dfm}

uses DM,Calendar, New_Bill_Tab;

//keyboard
//======================================
{----- собственно сами процедури и функции -----}

{получаем название раскладки}
function NameKeyboardLayout(layout : LongWord) : string;
var
  i: integer;
begin
  Result:='';
  try
    for i:=1 to CNT_LAYOUT do
      if TKbdValue[i]=layout then Result:= TKbdDisplayNames[i];
  except
    Result:='';
  end;
end;
//**************** end of NameKeyboardLayot ***************************
{активная раскладка в своей программе}
function GetActiveKbdLayout : LongWord;
begin
  result:= GetKeyboardLayout(0) shr $10;
end;
//***************** end of GetActiveKbdLayot ****************************
{активная раскладка в активном окне}
function GetActiveKbdLayoutWnd : LongWord;
var
  hWindow,idProcess : THandle;
begin
  // получить handle активного окна чужой программы
  hWindow := GetForegroundWindow;
  // Получить идентификатор чужого процесса
  idProcess := GetWindowThreadProcessId(hWindow,nil);
  // Получить текущую раскладку в чужой программе
  Result:=(GetKeyboardLayout(idProcess) shr $10);
end;
//***************** end of GetActiveKbdLayotWnd **************************
{установить раскладку в своей программе}
procedure SetKbdLayout(kbLayout : LongWord);
var
  Layout: HKL;
begin
  // Получить ссылку на раскладку
  Layout:=LoadKeyboardLayout(PChar(IntToStr(kbLayout)), 0);
  // Переключить раскладку на русскую
  ActivateKeyboardLayout(Layout,KLF_ACTIVATE);
end;
//****************** end of SetKbdLayot **********************************
{установить раскладку в активном окне}
procedure SetLayoutActiveWnd(kbLayout : LongWord);
var
  Layout: HKL;
  hWindow{, idProcess} : THandle; // ION T: не используется
begin
  // получить handle активного окна чужой программы
  hWindow := GetForegroundWindow;
  // Получить ссылку на раскладку
  Layout:=LoadKeyboardLayout(PChar(IntToStr(kbLayout)), 0);
  // посылаем сообщение о смене раскладки
  sendMessage(hWindow,WM_INPUTLANGCHANGEREQUEST,1,Layout);
end;
//***************** end of SetLayoutActiveWnd *****************************


//======================================


procedure TForm9.Button1Click(Sender: TObject);
begin
Form9.Close;
end;

procedure TForm9.Button2Click(Sender: TObject);
//var
 // zamet : string;
begin
//Update базы
zamet :=  Form9.Memo1.Lines.Text;

Text_SQL := 'UPDATE Bill_New SET ';
Text_SQL :=Text_SQL +'Zam = :p1';
Text_SQL :=Text_SQL + '  WHERE id = :p2;';

Form9.FDQuery1.Active := false;
Form9.FDQuery1.SQL.Clear;
Form9.FDQuery1.SQL.Add(Text_SQL);
Form9.FDQuery1.Params.ParamByName('p1').Value :=zamet ;
Form9.FDQuery1.Params.ParamByName('p2').Value :=Form9.kkk ;
Form9.FDQuery1.ExecSQL;


   Text_SQL :='select * from Bill_New ';

DataModule1.FDQuery_osn.Active :=false;
DataModule1.FDQuery_osn.SQL.Clear;
DataModule1.FDQuery_osn.SQL.Add(Text_SQL);
DataModule1.FDQuery_osn.Active :=true;

   Form1.DBGrid2.Columns[0].Width := 30;
   Form1.DBGrid2.Columns[0].Alignment:=taCenter;
   Form1.DBGrid2.Columns[1].Width := 80;
   Form1.DBGrid2.Columns[2].Width := 80;
   Form1.DBGrid2.Columns[3].Width := 80;
   Form1.DBGrid2.Columns[3].Alignment:=taCenter;
   Form1.DBGrid2.Columns[4].Width := 50;
   Form1.DBGrid2.Columns[4].Alignment:=taCenter;
   Form1.DBGrid2.Columns[5].Width := 80;
   Form1.DBGrid2.Columns[6].Width := 80;
   Form1.DBGrid2.Columns[7].Width := 80;
   Form1.DBGrid2.Columns[8].Width := 80;
   Form1.DBGrid2.Columns[9].Width := 80;
   Form1.DBGrid2.Columns[10].Width := 80;

   DataModule1.DataSource_osn.DataSet.Locate('Data',Form1.today_str,[loCaseInsensitive, loPartialKey]);

//Form1.FormCreate(Sender);



   Form9.Close;


end;

procedure TForm9.FormActivate(Sender: TObject);
begin
Form9.Memo1.SetFocus  ;
{
kkk := Form1.sel_row ;
Text_sql := 'select * from Bill_New where id = :p1';
Form9.FDQuery1.Active := false;
Form9.FDQuery1.SQL.Clear;
Form9.FDQuery1.SQL.Add(Text_SQL);
Form9.FDQuery1.Params.ParamByName('p1').Value :=kkk ;
Form9.FDQuery1.Active :=true;

zamet := Form9.FDQuery1.FieldByName('Zam').AsString;
Form9.Memo1.Text := zamet;
 s := 'Изменение заметок за  '
   + Form9.FDQuery1.FieldByName('Data').AsString
   + '  число';
Form9.Caption := s;
Form9.Memo1.SetFocus  ;
 }
end;

procedure TForm9.FormClose(Sender: TObject; var Action: TCloseAction);
begin
Form1.Show;
end;

procedure TForm9.FormShow(Sender: TObject);
begin
//
    LoadKeyboardLayout(StrCopy(Layout,'00000409'),KLF_ACTIVATE);
  {активная раскладка в своей программе}
    Label48.Caption:= NameKeyboardLayout(GetActiveKbdLayout);

end;

procedure TForm9.Label48Click(Sender: TObject);
begin
//
  if Form1.sortFlag1 then
  //На английский
  begin
    LoadKeyboardLayout(StrCopy(Layout,'00000409'),KLF_ACTIVATE);
    Label48.Caption:= NameKeyboardLayout(GetActiveKbdLayout);
  end
  else
  begin
    //На русский
    LoadKeyboardLayout( StrCopy(Layout,'00000419'),KLF_ACTIVATE);
    Label48.Caption:= NameKeyboardLayout(GetActiveKbdLayout);
  end;


Form1.sortFlag1 := not Form1.sortFlag1;

end;

end.
