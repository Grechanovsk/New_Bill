unit New_Bill_Tab;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs,
  //for inifiles
  inifiles,
  // для кодировки Url
  NetEncoding,
  Vcl.StdCtrls,
  // для использования временных утилит
  DateUtils,
  // для использвания ShellExecute
  ShellAPI,
  // чтобы не было хинтов
  FireDAC.Stan.Param, System.UITypes,
  // Для ini-файлов
  //IniFiles,
  //
  strutils, types,
  //для регулярных выражений
  System.RegularExpressionsCore,
  //использование внешних модулей
  DelphiNotesAge,
  UnGridToClb,    //использование модуля копирования
                   //DBGrid ClipBoard
  uMyExcel,   //функции Excel
  //для выбора дирректории
  FileCtrl,
  //для работы с Excel
  ComObj,
  ActiveX,
  //для Clipboard
  ClipBrd,
  //для вытягивания файла из инета
  Wininet,
  //
  Vcl.ExtCtrls, Data.DB, Vcl.ComCtrls, Vcl.Grids, Vcl.DBGrids, Vcl.Samples.Spin,
  Vcl.Menus, IdBaseComponent, IdComponent, IdTCPConnection, IdTCPClient, IdHTTP,
  Vcl.WinXCalendars, Vcl.DdeMan, Vcl.DBCtrls;


type
  TForm1 = class(TForm)
    Label30: TLabel;
    Edit10: TEdit;
    Timer1: TTimer;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    Panel1: TPanel;
    Label1: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label17: TLabel;
    Label16: TLabel;
    SpinEdit1: TSpinEdit;
    Edit7: TEdit;
    Edit8: TEdit;
    Panel2: TPanel;
    Label2: TLabel;
    Label5: TLabel;
    Label29: TLabel;
    SpinEdit2: TSpinEdit;
    Button3: TButton;
    SpinEdit3: TSpinEdit;
    Edit9: TEdit;
    Panel4: TPanel;
    Label6: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    Label15: TLabel;
    DBGrid2: TDBGrid;
    Edit1: TEdit;
    ComboBox1: TComboBox;
    ComboBox2: TComboBox;
    Edit2: TEdit;
    Edit3: TEdit;
    Edit4: TEdit;
    Edit5: TEdit;
    Edit6: TEdit;
    Button7: TButton;
    Button5: TButton;
    Panel3: TPanel;
    Panel5: TPanel;
    DBGrid1: TDBGrid;
    Button_opl: TButton;
    TabSheet4: TTabSheet;
    Panel10: TPanel;
    Label21: TLabel;
    Label20: TLabel;
    Panel11: TPanel;
    Label22: TLabel;
    Label23: TLabel;
    ComboBox4: TComboBox;
    Panel12: TPanel;
    Label24: TLabel;
    Label25: TLabel;
    Button16: TButton;
    Panel14: TPanel;
    Label26: TLabel;
    Label28: TLabel;
    Button20: TButton;
    ComboBox6: TComboBox;
    Button21: TButton;
    TabSheet2: TTabSheet;
    Panel7: TPanel;
    Button6: TButton;
    Memo1: TMemo;
    Panel9: TPanel;
    TabSheet3: TTabSheet;
    Panel6: TPanel;
    Label18: TLabel;
    Label19: TLabel;
    Button2: TButton;
    DateTimePicker1: TDateTimePicker;
    DateTimePicker2: TDateTimePicker;
    Button10: TButton;
    Button11: TButton;
    Button12: TButton;
    Button13: TButton;
    DBGrid3: TDBGrid;
    TabSheet5: TTabSheet;
    Panel13: TPanel;
    Label27: TLabel;
    Button17: TButton;
    Button18: TButton;
    Button19: TButton;
    TabSheet6: TTabSheet;
    Panel15: TPanel;
    Label31: TLabel;
    Label43: TLabel;
    DBGrid4: TDBGrid;
    Button22: TButton;
    DBGrid8: TDBGrid;
    Button27: TButton;
    Panel16: TPanel;
    Label32: TLabel;
    DBGrid5: TDBGrid;
    Button23: TButton;
    TabSheet7: TTabSheet;
    Panel8: TPanel;
    Label33: TLabel;
    Label41: TLabel;
    DBGrid6: TDBGrid;
    Button26: TButton;
    Edit12: TEdit;
    Panel17: TPanel;
    Label34: TLabel;
    Label35: TLabel;
    Label37: TLabel;
    Label38: TLabel;
    Label40: TLabel;
    DateTimePicker3: TDateTimePicker;
    DateTimePicker4: TDateTimePicker;
    Memo3: TMemo;
    Memo4: TMemo;
    Edit11: TEdit;
    Button24: TButton;
    Edit13: TEdit;
    TabSheet8: TTabSheet;
    Panel19: TPanel;
    Label36: TLabel;
    DBGrid7: TDBGrid;
    PopupMenu1: TPopupMenu;
    Menu11: TMenuItem;
    EXCEL1: TMenuItem;
    Button1: TButton;
    OpenDialog1: TOpenDialog;
    IdHTTP1: TIdHTTP;
    ClipBoard1: TMenuItem;
    SelectAll1: TMenuItem;
    Button28: TButton;
    Label39: TLabel;
    Button30: TButton;
    ComboBoxEx1: TComboBoxEx;
    ComboBoxEx2: TComboBoxEx;
    Button31: TButton;
    TabSheet9: TTabSheet;
    Button32: TButton;
    Panel20: TPanel;
    Label42: TLabel;
    Label44: TLabel;
    ComboBoxEx3: TComboBoxEx;
    Edit14: TEdit;
    Edit15: TEdit;
    SaveDialog1: TSaveDialog;
    Button34: TButton;
    Edit16: TEdit;
    Edit17: TEdit;
    Label45: TLabel;
    Edit18: TEdit;
    Edit19: TEdit;
    DBGrid9: TDBGrid;
    Timer2: TTimer;
    Label47: TLabel;
    Label48: TLabel;
    Label49: TLabel;
    TabSheet10: TTabSheet;
    Panel21: TPanel;
    ComboBoxEx4: TComboBoxEx;
    Label50: TLabel;
    Button15: TButton;
    Edit20: TEdit;
    Label51: TLabel;
    Label52: TLabel;
    SpinEdit4: TSpinEdit;
    Button9: TButton;
    Button8: TButton;
    Button4: TButton;
    Button25: TButton;
    Timer3: TTimer;
    Label53: TLabel;
    Button29: TButton;
    Button35: TButton;
    TabSheet11: TTabSheet;
    Button_Dengi: TButton;
    PopupMenu2: TPopupMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    Label55: TLabel;
    PageControl2: TPageControl;
    TabSheet12: TTabSheet;
    TabSheet13: TTabSheet;
    Label54: TLabel;
    Label56: TLabel;
    Memo6: TMemo;
    Panel18: TPanel;
    Memo7: TMemo;
    Button_Och: TButton;
    DBGrid11: TDBGrid;
    Label57: TLabel;
    Label58: TLabel;
    DBGrid10: TDBGrid;
    ComboBox3: TComboBox;
    Label59: TLabel;
    Button36: TButton;
    Button37: TButton;
    Label60: TLabel;
    ComboBox5: TComboBox;
    Label61: TLabel;
    ComboBox7: TComboBox;
    Label62: TLabel;
    ComboBox8: TComboBox;
    Label63: TLabel;
    PopupMenu_Govor: TPopupMenu;
    N3: TMenuItem;
    N4: TMenuItem;
    Button38: TButton;
    DateTimePicker5: TDateTimePicker;
    Label64: TLabel;
    Edit21: TEdit;
    Label65: TLabel;
    CheckBox1: TCheckBox;
    CheckBox2: TCheckBox;
    CheckBox3: TCheckBox;
    Button39: TButton;
    SpinEdit5: TSpinEdit;
    Label66: TLabel;
    RadioGroup1: TRadioGroup;
    RadioGroup2: TRadioGroup;
    Label67: TLabel;
    ComboBox9: TComboBox;
    MainMenu1: TMainMenu;
    N5: TMenuItem;
    N6: TMenuItem;
    N7: TMenuItem;
    N8: TMenuItem;
    N9: TMenuItem;
    N10: TMenuItem;
    N11: TMenuItem;
    Label46: TLabel;
    N12: TMenuItem;
    Paint1: TMenuItem;
    NotePad1: TMenuItem;
    Notepd1: TMenuItem;
    N13: TMenuItem;
    N14: TMenuItem;
    Button14: TButton;
    Button33: TButton;
    ComboBox10: TComboBox;
    Label68: TLabel;
    Label69: TLabel;
    ComboBox11: TComboBox;
    Button40: TButton;
    Button41: TButton;
    procedure Timer1Timer(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Edit1Change(Sender: TObject);
    procedure DBGrid2DrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure Button_oplClick(Sender: TObject);
    procedure Button7Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure Button16Click(Sender: TObject);
    procedure Button20Click(Sender: TObject);
    procedure Button21Click(Sender: TObject);
    procedure Button6Click(Sender: TObject);
    procedure TabSheet3Show(Sender: TObject);
    procedure Button10Click(Sender: TObject);
    procedure Button11Click(Sender: TObject);
    procedure Button12Click(Sender: TObject);
    procedure Button13Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button17Click(Sender: TObject);
    procedure Button18Click(Sender: TObject);
    procedure Button19Click(Sender: TObject);
    procedure DBGrid4DrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure DBGrid5DrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure DBGrid4DblClick(Sender: TObject);
    procedure DBGrid5DblClick(Sender: TObject);
    procedure Button22Click(Sender: TObject);
    procedure Button23Click(Sender: TObject);
    procedure TabSheet7Show(Sender: TObject);
    procedure DBGrid6DrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure Button24Click(Sender: TObject);
    procedure Button26Click(Sender: TObject);
    procedure Menu11Click(Sender: TObject);
    procedure EXCEL1Click(Sender: TObject);
    procedure ClipBoard1Click(Sender: TObject);
    procedure SelectAll1Click(Sender: TObject);
    procedure Button28Click(Sender: TObject);
    procedure DBGrid3DrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure TabSheet6Show(Sender: TObject);
    procedure DBGrid8DrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure Button27Click(Sender: TObject);
    procedure EditFor10Click(Sender: TObject);
    procedure DBGrid2DblClick(Sender: TObject);
    procedure DBGrid7DblClick(Sender: TObject);
    procedure Button30Click(Sender: TObject);
    procedure TabSheet4Show(Sender: TObject);
    procedure ComboBoxEx2KeyPress(Sender: TObject; var Key: Char);
    procedure Button31Click(Sender: TObject);
    procedure Button32Click(Sender: TObject);
    procedure Button34Click(Sender: TObject);
    procedure StringGrid1DrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
    procedure Edit16Change(Sender: TObject);
    procedure DBGrid9TitleClick(Column: TColumn);
    procedure Edit17Change(Sender: TObject);
    procedure DBGrid9DrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure Timer2Timer(Sender: TObject);
    procedure Label48Click(Sender: TObject);
    procedure ComboBoxEx1Select(Sender: TObject);
    procedure ComboBox4Select(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure ComboBoxEx4Select(Sender: TObject);
    procedure TabSheet10Show(Sender: TObject);
    procedure Button15Click(Sender: TObject);
    procedure Edit20Click(Sender: TObject);
    procedure Edit17Click(Sender: TObject);
//    procedure ComboBoxEx8Select(Sender: TObject);
    procedure ComboBoxEx1Change(Sender: TObject);
    procedure Button8Click(Sender: TObject);
    procedure Button9Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Button25Click(Sender: TObject);
    procedure DBGrid9CellClick(Column: TColumn);
    procedure DBGrid5CellClick(Column: TColumn);
    procedure ShowTime;
    procedure Timer3Timer(Sender: TObject);
    procedure Button29Click(Sender: TObject);
    procedure Button35Click(Sender: TObject);
    procedure Button_DengiClick(Sender: TObject);
    procedure TabSheet11Show(Sender: TObject);
    procedure N1Click(Sender: TObject);
    procedure N2Click(Sender: TObject);
    procedure TabSheet13Show(Sender: TObject);
    procedure TabSheet12Show(Sender: TObject);
    procedure Button_OchClick(Sender: TObject);
    procedure ComboBox3Select(Sender: TObject);
    procedure Button37Click(Sender: TObject);
    procedure Button36Click(Sender: TObject);
    procedure ComboBox5Select(Sender: TObject);
    procedure ComboBox7Select(Sender: TObject);
    procedure ComboBox8Select(Sender: TObject);
    procedure N3Click(Sender: TObject);
    procedure N4Click(Sender: TObject);
    procedure Button38Click(Sender: TObject);
    procedure Button39Click(Sender: TObject);
    procedure DateTimePicker5Change(Sender: TObject);
    procedure SpinEdit5Change(Sender: TObject);
    procedure SpinEdit5Click(Sender: TObject);
    procedure ComboBox9Select(Sender: TObject);
    procedure N11Click(Sender: TObject);
    procedure N8Click(Sender: TObject);
    procedure N9Click(Sender: TObject);
    procedure N12Click(Sender: TObject);
    procedure Paint1Click(Sender: TObject);
    procedure NotePad1Click(Sender: TObject);
    procedure Notepd1Click(Sender: TObject);
    procedure N13Click(Sender: TObject);
    procedure N14Click(Sender: TObject);
    procedure N6Click(Sender: TObject);
    procedure Button14Click(Sender: TObject);
    procedure Button33Click(Sender: TObject);
    procedure ComboBox10Select(Sender: TObject);
    procedure ComboBox11Select(Sender: TObject);
    procedure Button40Click(Sender: TObject);
    procedure Button41Click(Sender: TObject);

//    procedure FitGrid(Grid: TDBGrid);
  private
    { Private declarations }
  public
    { Public declarations }
    tm_work  :  integer;
    today_str : string;

    ejem_platej     :integer;

   day             :integer;
   mes             :integer;
   god             :integer;

   Text_label2     :string;

   sel_row :      integer;
   sortFlag:      Boolean;
   sortFlag1:      Boolean;
   Layout: array[0.. KL_NAMELENGTH] of char;
   //For ini file
   Backup_drive  :string;
   Database_path :string;
   comp_name     :string;
   backup_name   :string;
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
   {var
  Months: array[1..12] of String = ('января', 'февраля', 'марта',
                                         'апреля', 'мая', 'июня',
                                         'июля', 'августа', 'сентября',
                                         'октября', 'ноября', 'декабря');}
  Months:array [1..12] of string[8]=('Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь', 'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь');
  //  ShowMessage('Сейчас '+Months[MonthOf(Now)]+' '+IntToStr(YearOf(Now))+'  года')
var
  Form1: TForm1;

implementation

{$R *.dfm}

uses DM, Bill2, Bill3, Bill_Opl_Sveta, Bill_Opl_Sheta, Bill_Svet, Bill_Shet,
  Bill_Pokz_Voda, Edit_Zam, Bill_Login_Sel, Calendar, Nstroyki;
//===========================================================================
function ClearDir( Dir: string ): boolean;
var
  isFound: boolean;
  sRec: TSearchRec;
begin
   Result := false;
   ChDir( Dir );
   if IOResult <> 0 then
   begin
      ShowMessage( 'Не могу войти в каталог: ' + Dir );
      Exit;
   end;
   if Dir[Length(Dir)] <> '\' then Dir := Dir + '\';
   isFound := FindFirst( Dir + '*.*', faAnyFile, sRec ) = 0;
   while isFound do
   begin
   if ( sRec.Name <> '.' ) and ( sRec.Name <> '..' ) then
      if ( sRec.Attr and faDirectory ) = faDirectory then
      begin
         if not ClearDir( Dir + sRec.Name ) then
            Exit;
         if ( sRec.Name <> '.' ) and ( sRec.Name <> '..' ) then
            if ( Dir + sRec.Name ) <> Dir then
            begin
               ChDir( '..' );
               RmDir( Dir + sRec.Name );
            end;
      end
   else
         if not DeleteFile( Dir + sRec.Name ) then
         begin
            ShowMessage( 'Не могу удалить файл: ' + sRec.Name );
            Exit;
         end;
      isFound := FindNext( sRec ) = 0;
   end;
   FindClose( sRec );
   Result := IOResult = 0;
end;
//===========================================================================
procedure Perehod (Perehod : string);
begin
 ShellExecute(Application.Handle, 'open',PChar(Perehod), nil, nil, SW_SHOWNORMAL) ;
end;
//============================================================================
procedure ExecAndWait (CommandLine : string);
var
  Rlst                  : LongBool; //результат выполнения
  StartUpInfo           : TStartUpInfo; //параметры будущего процесса
  ProcessInfo           : TProcessInformation; //Отслеживание выполнения
  ExitCode              : Cardinal; //код завершения
  Error                 : integer; //номер ошибкок
 begin
  //Заполнение нулями всего StartUpInfo.
   FillChar(StartUpInfo, SizeOf(TStartUpInfo), 0);
  //После этого выставляем в нем некоторые параметры.
  with StartUpInfo do
   begin
   //Содержит количество байтов, занимаемых структурой TStartUpInfo.
   //Обязательно для заполнения. Инициализируйте как SizeOf(TStartUpInfo).
    cb := SizeOf(TStartUpInfo);
   //Содержит набор флагов, позволяющих управлять созданием дочернего процесса.
   //Показываем окно, курсор - часики.
    dwFlags := STARTF_USESHOWWINDOW or STARTF_FORCEONFEEDBACK;
   //Определяет как должно выглядеть окно запущенного приложения.
   //Нормальное отображение
    wShowWindow := SW_SHOWNORMAL;
   end;
    Rlst:= CreateProcess(nil, PChar(CommandLine),
    nil, nil, false, NORMAL_PRIORITY_CLASS,
    nil, nil, StartUpInfo, ProcessInfo);

  //Отслеживаем выполнение.
  //Если запуск успешен
  if Rlst then
   with ProcessInfo do
     begin
      //Ждем завершения инициализации.
      WaitForInputIdle(hProcess, INFINITE);
      //Ждем завершения процесса.
      WaitforSingleObject(ProcessInfo.hProcess, INFINITE);
      //Получаем код завершения.
      GetExitCodeProcess(ProcessInfo.hProcess, ExitCode);
      //Закрываем дескриптор процесса.
      CloseHandle(hThread);
      //Закрываем дескриптор потока.
      CloseHandle(hProcess);
     end

   else
      begin
      //В случае ошибки - выводим на экран сообщение.
       Error := GetLastError;
       MessageDlg(SysErrorMessage(Error), mtError, [mbOk], 0);
      end;

 end;

//keyboard
//======================================
{----- собственно сами процедури и функции -----}
    function ReadComputerName: string;
    var
      Size: Cardinal;
    begin
      Size := MAX_COMPUTERNAME_LENGTH + 1;
      SetLength(Result, Size);
      GetComputerName(PChar(Result), Size);
      // Урезаем строку до действительной длины имени компьютера
      SetLength(Result, Size);
    end;

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
procedure TForm1.ShowTime;
var Time : TDateTime; // текущее время
begin
 Time := Now(); // получить системное время
 Label53.Caption := FormatDateTime('hh:mm:ss',Time);
end;




procedure TForm1.SpinEdit5Change(Sender: TObject);
begin
 Form1.DateTimePicker5.Enabled := false;
end;

procedure TForm1.SpinEdit5Click(Sender: TObject);
begin
Form1.SpinEdit5.Text := '';
end;

//======================================
function GetDays(ADate: TDate): Extended;
 var
   FirstOfYear: TDateTime;
 begin
   FirstOfYear := EncodeDate(StrToInt(FormatDateTime('yyyy', now)) - 1, 12, 31);
   Result      := ADate - FirstOfYear;
 end;
//======================================
procedure FitGrid(Grid: TDBGrid);
const
  C_Add=3;
var
  ds: TDataSet;
  bm: TBookmark;
  i: Integer;
  w: Integer;
  a: Array of Integer;
begin
  ds := Grid.DataSource.DataSet;
  if Assigned(ds) then
  begin
    ds.DisableControls;
    bm := ds.GetBookmark;
    try
      ds.First;
      SetLength(a, Grid.Columns.Count);
      while not ds.Eof do
      begin
        for I := 0 to Grid.Columns.Count - 1 do
        begin
          if Assigned(Grid.Columns[i].Field) then
          begin
            w :=  Grid.Canvas.TextWidth(ds.FieldByName(Grid.Columns[i].Field.FieldName).DisplayText);
            if a[i] < w  then
               a[i] := w ;
          end;
        end;
        ds.Next;
      end;
      //if fieldwidth is smaller than Row 0 (field names) fix
      for I := 0 to Grid.Columns.Count - 1 do
      begin
        w :=  Grid.Canvas.TextWidth(Grid.Columns[i].Field.FieldName);
        if a[i] < w  then
           a[i] := w ;
      end;

      for I := 0 to Grid.Columns.Count - 1 do
        Grid.Columns[i].Width := a[i] + C_Add;
        ds.GotoBookmark(bm);
    finally
      ds.FreeBookmark(bm);
      ds.EnableControls;
    end;
  end;
end;
//======================================

//=====================================
function Tokenize(Str: WideString; Delimiter: string): TStringList;
var
  tmpStrList: TStringList;
  tmpString, tmpVal: WideString;
  DelimPos: LongInt;
begin
  tmpStrList := TStringList.Create;
  TmpString := Str;
  DelimPos := 1;
  while DelimPos > 0 do
  begin
    DelimPos := LastDelimiter(Delimiter, TmpString);
    tmpVal := Copy(TmpString, DelimPos + 1, Length(TmpString));
    if tmpVal <> '' then
      tmpStrList.Add(UpperCase(tmpVal));
    Delete(TmpString, DelimPos, Length(TmpString));
  end;
  Tokenize := tmpStrList;
end;
//=====================================

function GetInetFile(const fileURL, FileName: string): boolean;
const
  BufferSize = 1024;
var
  hSession, hURL: HInternet;
  Buffer: array[1..BufferSize] of Byte;
  BufferLen: DWORD;
  f: file;
  sAppName: string;
begin
  Result:=False;
  sAppName := ExtractFileName(Application.ExeName);
  hSession := InternetOpen(PChar(sAppName),
  INTERNET_OPEN_TYPE_PRECONFIG,
  nil, nil, 0);
  try
    hURL := InternetOpenURL(hSession,
    PChar(fileURL),
    nil,0,0,0);
    try
      AssignFile(f, FileName);
      Rewrite(f,1);
      repeat
        InternetReadFile(hURL, @Buffer,
        SizeOf(Buffer), BufferLen);
        BlockWrite(f, Buffer, BufferLen)
      until
        BufferLen = 0;
      CloseFile(f);
      Result:=True;
    finally
      InternetCloseHandle(hURL)
    end;
  finally
    InternetCloseHandle(hSession)
  end;
end;


{ Удалить каталог со всем содержимым }
function DeleteDir(Dir : string) : boolean;
Var
Found : integer;
SearchRec : TSearchRec;
begin
result:=false;
if IOResult<>0 then ;
ChDir(Dir);
if IOResult<>0 then begin
ShowMessage('Не могу войти в каталог: '+Dir); exit;
end;
Found := FindFirst('*.*', faAnyFile, SearchRec);
while Found = 0 do
begin
if (SearchRec.Name<>'.')and(SearchRec.Name<>'..') then
if (SearchRec.Attr and faDirectory)<>0 then begin
if not DeleteDir(SearchRec.Name) then exit;
end else
if not DeleteFile(SearchRec.Name) then begin
ShowMessage('Не могу удалить файл: '+SearchRec.Name); exit;
end;
Found := FindNext(SearchRec);
end;
FindClose(SearchRec);
ChDir('..'); RmDir(Dir);
result:=IOResult=0;
end;

function CopyDir(const fromDir, toDir: string): Boolean;
var
  fos: TSHFileOpStruct;
begin
  ZeroMemory(@fos, SizeOf(fos));
  with fos do
  begin
    wFunc  := FO_COPY;
    fFlags := FOF_FILESONLY;
    pFrom  := PChar(fromDir + #0);
    pTo    := PChar(toDir)
  end;
  Result := (0 = ShFileOperation(fos));
end;

function GridSelectAll(Grid: TDBGrid): Longint;
begin
  Result := 0;
  Grid.SelectedRows.Clear;
  with Grid.Datasource.DataSet do
  begin
    First;
    DisableControls;
    try
      while not EOF do
      begin
        Grid.SelectedRows.CurrentRowSelected := True;
        inc(Result);
        Next;
      end;
    finally
      EnableControls;
    end;
  end;
end;

function ParseStr(s, tag1, tag2: string): string;
var
  S1: string;
  S2: string;
 // S3: string;
 // S4: string;
  Len1 : integer;
  Len2 : integer;
  pos1 : integer;
  pos2 : integer;
 begin
//   s :=  Utf8ToAnsi(s);
   len1 := Length(tag1);
   len2 := Length(tag2);
   pos1 := AnsiPos(tag1,S);
   S1 := copy(s,pos1+len1,Length(S)-(pos1+len1));
   pos1 := AnsiPos(tag1,S);
   pos2 := AnsiPos(tag2,S1);
   S2 := Copy(S1,1,pos2-1);

     Result:= S2   ;
end;


procedure CopySelGRDToClb(dbg: TDBGrid);
var
  i,j: Integer;
  S: String;
begin
 S := '';
 dbg.DataSource.DataSet.DisableControls;
 Clipboard.Open;
 for i := 0 to dbg.SelectedRows.Count - 1
 do begin
    dbg.DataSource.DataSet.GotoBookmark(dbg.SelectedRows.Items[i]);
    for j:= 0 to dbg.DataSource.DataSet.FieldCount - 1
    do S := S + #9 + dbg.DataSource.DataSet.Fields.Fields[j].AsString;
    S := S + #13;
    end;
 Clipboard.AsText := S;
 Clipboard.Close;
 dbg.DataSource.DataSet.EnableControls;
end;

function DateToBase(ish_str: string): string;
begin
Result  := Copy(ish_str,7,4)+ '-'+Copy(ish_str,4,2)+'-'+Copy(ish_str,1,2);
end;
{
function CheckExcelInstall:boolean;
var
  ClassID: TCLSID;
  Rez : HRESULT;
begin
// Ищем CLSID OLE-объекта
  Rez := CLSIDFromProgID(PWideChar(WideString('Excel.Application')), ClassID);
  if Rez = S_OK then  // Объект найден
    Result := true
  else
    Result := false;
end;
 }


procedure TForm1.Button10Click(Sender: TObject);
begin
Form1.DateTimePicker1.Date := IncMonth(Form1.DateTimePicker1.Date);
end;

procedure TForm1.Button11Click(Sender: TObject);
var
 wy, wm, wd :Word;
begin
if Form1.DateTimePicker1.Date > StrToDate(Trawm_ish)  then
   begin
    DecodeDate( Form1.DateTimePicker1.Date, wy, wm, wd );
    Dec( wm );
    if (wm = 0 ) then
      begin
      wm := 12;
      Dec( wy );
      end;
   Form1.DateTimePicker1.Date := EncodeDate( wy, wm, wd );
   end
 else
  begin
    Form1.DateTimePicker1.Date := StrToDate(Trawm_ish);
  end;

end;

procedure TForm1.Button12Click(Sender: TObject);
begin
Form1.DateTimePicker2.Date := IncMonth(Form1.DateTimePicker2.Date);
end;

procedure TForm1.Button13Click(Sender: TObject);
var
 wy, wm, wd :Word;
begin
if Form1.DateTimePicker2.Date > StrToDate(Trawm_ish)  then
   begin
    DecodeDate( Form1.DateTimePicker2.Date, wy, wm, wd );
    Dec( wm );
    if (wm = 0 ) then
      begin
      wm := 12;
      Dec( wy );
      end;
   Form1.DateTimePicker2.Date := EncodeDate( wy, wm, wd );
   end
 else
  begin
    Form1.DateTimePicker2.Date := StrToDate(Trawm_ish);
  end;

end;



procedure TForm1.Button14Click(Sender: TObject);
var
 ApN :string;
begin
ApN := PChar(Application.ExeName) ;
ShellExecute(Handle,'open',PChar(ApN), nil, nil, SW_SHOWNORMAL) ;
Application.Terminate;
end;

procedure TForm1.Button15Click(Sender: TObject);
var
  Ini: Tinifile; //необходимо создать объект, чтоб потом с ним работатьbegin
begin
//Database_path
//создали файл в директории программы
  Ini:=TiniFile.Create(extractfilepath(paramstr(0))+'My_Ini.ini');

  if Form1.ComboBoxEx4.Text = '' then
  begin
  // Form1.ComboBoxEx4.Text := 'D:\';
  ShowMessage('Ты не установил диск для архивов перейди на вкладку Set, и установи все нужные значения.');
  Form1.PageControl1.ActivePage := TabSheet10;
  exit;
  end
else
  begin
  Ini.WriteString('Main','Drive_Backup',Form1.Backup_drive);
  end;
//---------------------------------------------------
  if Form1.Edit20.Text = '' then
  begin
  // Form1.ComboBoxEx4.Text := 'D:\';
  ShowMessage('Ты не установил путь для рабочей DB перейди на вкладку Set, и установи все нужные значения.');
  Form1.PageControl1.ActivePage := TabSheet10;
  exit;
  end
else
  begin
  Ini.WriteString('Main','Database_path',
    Form1.Edit20.Text);
  end;
//-------------------------------------------

  if Form1.SpinEdit4.Value = 0 then
  begin
  // Form1.ComboBoxEx4.Text := 'D:\';
  ShowMessage('Ты не установил ежемесячную оплату.');
  Form1.PageControl1.ActivePage := TabSheet10;
  exit;
  end
else
  begin
  Ini.WriteInteger('Main','Ежемесячный платёж',Form1.SpinEdit4.Value);
  end;


 // Ini.WriteInteger('Size','Width',form1.width);
 // Ini.WriteInteger('Size','Height',form1.height);
 // Ini.WriteInteger('Position','X',form1.left);
 // Ini.WriteInteger('Position','Y',form1.top);
  Ini.Free;
//==============================================
//перечитываем сохраненные настройки
//открываем файл
  Ini:=TiniFile.Create(extractfilepath(paramstr(0))+'My_Ini.ini');
  Backup_drive := Ini.ReadString('Main','Drive_Backup',Form1.ComboBoxEx1.Text);
  Database_path := Ini.ReadString('Main','Database_path',Form1.Edit20.Text);
  ejem_platej :=   Ini.ReadInteger('Main','Ежемесячный платёж',Form1.SpinEdit4.Value);
//  Form1.Width:=Ini.ReadInteger('Size','Width',100);
  //последнее значение (100) это значение по умолчанию (default)
//  Form1.Height:=Ini.ReadInteger('Size','Height',100);
//  Form1.Left:=Ini.ReadInteger('Position','X',10);
//  Form1.Top:=Ini.ReadInteger('Position','Y',10);
  Ini.Free;
  Form1.Button14Click(Sender);

end;

procedure TForm1.Button16Click(Sender: TObject);
var
it       : integer;
begin
it := Form1.ComboBoxEx2.ItemIndex;
case it of
//C:\Program Files\VideoLAN\VLC\vlc.exe
//Звезда
0: ShellExecute(Application.Handle, 'open', 'C:\Program Files\VideoLAN\VLC\vlc.exe',
  'http://unicast.netlux.org/media/zvezda', nil,SW_SHOWNOACTIVATE);
//Золотая коллекция
1:  ShellExecute(Application.Handle, 'open', 'C:\Program Files\VideoLAN\VLC\vlc.exe',
  'http://unicast.netlux.org/media/tdk', nil,SW_SHOWNOACTIVATE);
//Россия 1
2:  ShellExecute(Application.Handle, 'open', 'C:\Program Files\VideoLAN\VLC\vlc.exe',
  'http://unicast.netlux.org/media/rus1', nil,SW_SHOWNOACTIVATE);
//Россия 24
3:   ShellExecute(Application.Handle, 'open', 'C:\Program Files\VideoLAN\VLC\vlc.exe',
  'http://unicast.netlux.org/media/rus24', nil,SW_SHOWNOACTIVATE);
//5 канал
4:  ShellExecute(Application.Handle, 'open', 'C:\Program Files\VideoLAN\VLC\vlc.exe',
  'http://unicast.netlux.org/media/5ch', nil,SW_SHOWNOACTIVATE);
//Amedia
5:   ShellExecute(Application.Handle, 'open', 'C:\Program Files\VideoLAN\VLC\vlc.exe',
  'http://unicast.netlux.org/media/amedia', nil,SW_SHOWNOACTIVATE);
//Дом Кино Премиум HD
6: ShellExecute(Application.Handle, 'open', 'C:\Program Files\VideoLAN\VLC\vlc.exe',
  'http://unicast.netlux.org/media/domkinopre', nil,SW_SHOWNOACTIVATE);
//Дом Кино
7:   ShellExecute(Application.Handle, 'open', 'C:\Program Files\VideoLAN\VLC\vlc.exe',
  'http://unicast.netlux.org/media/domkino', nil,SW_SHOWNOACTIVATE);
//Известия
8:   ShellExecute(Application.Handle, 'open', 'C:\Program Files\VideoLAN\VLC\vlc.exe',
  'http://unicast.netlux.org/media/izvestiya', nil,SW_SHOWNOACTIVATE);
//Иллюзион
9:  ShellExecute(Application.Handle, 'open', 'C:\Program Files\VideoLAN\VLC\vlc.exe',
  'http://unicast.netlux.org/media/illuzplus', nil,SW_SHOWNOACTIVATE);
//Русский иллюзион
10:   ShellExecute(Application.Handle, 'open', 'C:\Program Files\VideoLAN\VLC\vlc.exe',
  'http://unicast.netlux.org/media/rusilluz', nil,SW_SHOWNOACTIVATE);
//Кинокомедия
11:   ShellExecute(Application.Handle, 'open', 'C:\Program Files\VideoLAN\VLC\vlc.exe',
  'http://unicast.netlux.org/media/kinokomediya', nil,SW_SHOWNOACTIVATE);
//Русская комедия
12:   ShellExecute(Application.Handle, 'open', 'C:\Program Files\VideoLAN\VLC\vlc.exe',
  'http://unicast.netlux.org/media/ruscomedia', nil,SW_SHOWNOACTIVATE);
//Киномикс
13:   ShellExecute(Application.Handle, 'open', 'C:\Program Files\VideoLAN\VLC\vlc.exe',
  'http://unicast.netlux.org/media/kinomix', nil,SW_SHOWNOACTIVATE);
//Киносвидание
14:   ShellExecute(Application.Handle, 'open', 'C:\Program Files\VideoLAN\VLC\vlc.exe',
  'http://unicast.netlux.org/media/kinosvid', nil,SW_SHOWNOACTIVATE);
//Киносемья
15:   ShellExecute(Application.Handle, 'open', 'C:\Program Files\VideoLAN\VLC\vlc.exe',
  'http://unicast.netlux.org/media/kinosemya', nil,SW_SHOWNOACTIVATE);
//Киносерия
16:   ShellExecute(Application.Handle, 'open', 'C:\Program Files\VideoLAN\VLC\vlc.exe',
  'http://unicast.netlux.org/media/kinoseriya', nil,SW_SHOWNOACTIVATE);
//Кинохит
17:   ShellExecute(Application.Handle, 'open', 'C:\Program Files\VideoLAN\VLC\vlc.exe',
  'http://unicast.netlux.org/media/kinohit', nil,SW_SHOWNOACTIVATE);
//Кто есть кто
18:   ShellExecute(Application.Handle, 'open', 'C:\Program Files\VideoLAN\VLC\vlc.exe',
  'http://unicast.netlux.org/media/ktoestkto', nil,SW_SHOWNOACTIVATE);
//Любимое кино
19:   ShellExecute(Application.Handle, 'open', 'C:\Program Files\VideoLAN\VLC\vlc.exe',
  'http://unicast.netlux.org/media/lubimoe', nil,SW_SHOWNOACTIVATE);
//Мир сериала
20:   ShellExecute(Application.Handle, 'open', 'C:\Program Files\VideoLAN\VLC\vlc.exe',
  'http://unicast.netlux.org/media/mirseriala', nil,SW_SHOWNOACTIVATE);
//Мужское кино
21:   ShellExecute(Application.Handle, 'open', 'C:\Program Files\VideoLAN\VLC\vlc.exe',
  'http://unicast.netlux.org/media/muzhskoekino', nil,SW_SHOWNOACTIVATE);
//Мужской
22:   ShellExecute(Application.Handle, 'open', 'C:\Program Files\VideoLAN\VLC\vlc.exe',
  'http://unicast.netlux.org/media/muzhskoy', nil,SW_SHOWNOACTIVATE);
//НТВ
23:   ShellExecute(Application.Handle, 'open', 'C:\Program Files\VideoLAN\VLC\vlc.exe',
  'http://unicast.netlux.org/media/ntv', nil,SW_SHOWNOACTIVATE);
//НТВ Право
24:   ShellExecute(Application.Handle, 'open', 'C:\Program Files\VideoLAN\VLC\vlc.exe',
  'http://unicast.netlux.org/media/ntvpravo', nil,SW_SHOWNOACTIVATE);
//НТВ Сериал
25:   ShellExecute(Application.Handle, 'open', 'C:\Program Files\VideoLAN\VLC\vlc.exe',
  'http://unicast.netlux.org/media/ntvserial', nil,SW_SHOWNOACTIVATE);
//Наше новое кино
26:   ShellExecute(Application.Handle, 'open', 'C:\Program Files\VideoLAN\VLC\vlc.exe',
  'http://unicast.netlux.org/media/nashenovoe', nil,SW_SHOWNOACTIVATE);
//Новороссия
27:   ShellExecute(Application.Handle, 'open', 'C:\Program Files\VideoLAN\VLC\vlc.exe',
  'http://unicast.netlux.org/media/novorossia', nil,SW_SHOWNOACTIVATE);
//Ностальгия
28:   ShellExecute(Application.Handle, 'open', 'C:\Program Files\VideoLAN\VLC\vlc.exe',
  'http://unicast.netlux.org/media/nostalgiya', nil,SW_SHOWNOACTIVATE);
//Оплот 2
29:   ShellExecute(Application.Handle, 'open', 'C:\Program Files\VideoLAN\VLC\vlc.exe',
  'http://unicast.netlux.org/media/oplot2', nil,SW_SHOWNOACTIVATE);
//Оружие
30:   ShellExecute(Application.Handle, 'open', 'C:\Program Files\VideoLAN\VLC\vlc.exe',
  'http://unicast.netlux.org/media/oruzhie', nil,SW_SHOWNOACTIVATE);
//Первый
31:   ShellExecute(Application.Handle, 'open', 'C:\Program Files\VideoLAN\VLC\vlc.exe',
  'http://unicast.netlux.org/media/perviy', nil,SW_SHOWNOACTIVATE);
//Первый космический
32:   ShellExecute(Application.Handle, 'open', 'C:\Program Files\VideoLAN\VLC\vlc.exe',
  'http://unicast.netlux.org/media/eureka', nil,SW_SHOWNOACTIVATE);
//Победа
33:   ShellExecute(Application.Handle, 'open', 'C:\Program Files\VideoLAN\VLC\vlc.exe',
  'http://unicast.netlux.org/media/pobeda', nil,SW_SHOWNOACTIVATE);
// Пятница
34:   ShellExecute(Application.Handle, 'open', 'C:\Program Files\VideoLAN\VLC\vlc.exe',
  'http://unicast.netlux.org/media/friday', nil,SW_SHOWNOACTIVATE);
//Родное кино
35:   ShellExecute(Application.Handle, 'open', 'C:\Program Files\VideoLAN\VLC\vlc.exe',
  'http://unicast.netlux.org/media/rodnoe', nil,SW_SHOWNOACTIVATE);
//Русская комедия
36:   ShellExecute(Application.Handle, 'open', 'C:\Program Files\VideoLAN\VLC\vlc.exe',
  'http://unicast.netlux.org/media/ruscomedia', nil,SW_SHOWNOACTIVATE);
//Русский бестселлер
37:   ShellExecute(Application.Handle, 'open', 'C:\Program Files\VideoLAN\VLC\vlc.exe',
  'http://unicast.netlux.org/media/rusbestseller', nil,SW_SHOWNOACTIVATE);
//Русский детектив
38:   ShellExecute(Application.Handle, 'open', 'C:\Program Files\VideoLAN\VLC\vlc.exe',
  'http://unicast.netlux.org/media/rusdetektiv', nil,SW_SHOWNOACTIVATE);
//Русский роман
39:   ShellExecute(Application.Handle, 'open', 'C:\Program Files\VideoLAN\VLC\vlc.exe',
  'http://unicast.netlux.org/media/rusroman', nil,SW_SHOWNOACTIVATE);
// Спас
40:   ShellExecute(Application.Handle, 'open', 'C:\Program Files\VideoLAN\VLC\vlc.exe',
  'http://unicast.netlux.org/media/spas', nil,SW_SHOWNOACTIVATE);
//Суббота
41:   ShellExecute(Application.Handle, 'open', 'C:\Program Files\VideoLAN\VLC\vlc.exe',
  'http://unicast.netlux.org/media/super', nil,SW_SHOWNOACTIVATE);
//Т 24
42:   ShellExecute(Application.Handle, 'open', 'C:\Program Files\VideoLAN\VLC\vlc.exe',
  'http://unicast.netlux.org/media/t24', nil,SW_SHOWNOACTIVATE);
//ТВ 3
43:   ShellExecute(Application.Handle, 'open', 'C:\Program Files\VideoLAN\VLC\vlc.exe',
  'http://unicast.netlux.org/media/tv3', nil,SW_SHOWNOACTIVATE);
//Юнион
44:   ShellExecute(Application.Handle, 'open', 'C:\Program Files\VideoLAN\VLC\vlc.exe',
  'http://unicast.netlux.org/media/union', nil,SW_SHOWNOACTIVATE);
//Discovery Channel
45:   ShellExecute(Application.Handle, 'open', 'C:\Program Files\VideoLAN\VLC\vlc.exe',
  'http://unicast.netlux.org/media/discovery', nil,SW_SHOWNOACTIVATE);
//Discovery Science
46:   ShellExecute(Application.Handle, 'open', 'C:\Program Files\VideoLAN\VLC\vlc.exe',
  'http://unicast.netlux.org/media/discovery_science_test', nil,SW_SHOWNOACTIVATE);
//National Geographic HD
47:   ShellExecute(Application.Handle, 'open', 'C:\Program Files\VideoLAN\VLC\vlc.exe',
  'http://unicast.netlux.org/media/natgeohd', nil,SW_SHOWNOACTIVATE);
//Ren TV
48:   ShellExecute(Application.Handle, 'open', 'C:\Program Files\VideoLAN\VLC\vlc.exe',
  'http://unicast.netlux.org/media/rentv', nil,SW_SHOWNOACTIVATE);
//Viasat History
49:   ShellExecute(Application.Handle, 'open', 'C:\Program Files\VideoLAN\VLC\vlc.exe',
  'http://unicast.netlux.org/media/viahistory', nil,SW_SHOWNOACTIVATE);
//Viasat Nature/History HD
50:   ShellExecute(Application.Handle, 'open', 'C:\Program Files\VideoLAN\VLC\vlc.exe',
  'http://unicast.netlux.org/media/vianathis', nil,SW_SHOWNOACTIVATE);
//TV1000
51:   ShellExecute(Application.Handle, 'open', 'C:\Program Files\VideoLAN\VLC\vlc.exe',
  'http://unicast.netlux.org/media/tv1000', nil,SW_SHOWNOACTIVATE);
//TV1000 Action
52:   ShellExecute(Application.Handle, 'open', 'C:\Program Files\VideoLAN\VLC\vlc.exe',
  'http://unicast.netlux.org/media/tv1000act', nil,SW_SHOWNOACTIVATE);
//TV1000 Русское кино
53:   ShellExecute(Application.Handle, 'open', 'C:\Program Files\VideoLAN\VLC\vlc.exe',
  'http://unicast.netlux.org/media/tv1000rus', nil,SW_SHOWNOACTIVATE);
  Else
   ShellExecute(Application.Handle, 'open', 'C:\Program Files\VideoLAN\VLC\vlc.exe',
  'd:\unicast.m3u', nil,SW_SHOWNOACTIVATE);
end;

end;

procedure TForm1.Button17Click(Sender: TObject);
var
pinging : integer;
pinging_str : string;
begin
pinging := RadioGroup1.ItemIndex;

case pinging of
//
0:  pinging_str := '192.168.1.1';
//
1:   pinging_str := '8.8.8.8';
//
2:   pinging_str := 'ya.ru';
//
3:   pinging_str := 'nii_rzdpm.dnmu.ru';
//
4:   pinging_str := 'www.netlux.org';

  Else
   pinging_str := '192.168.1.1';
  end;
   pinging_str := '/k ping.exe -t ' + pinging_str;
ShellExecute(0,'open','cmd.exe',PWideChar(pinging_str),
'C:\Windows\system32\',SW_SHOW);
end;

procedure TForm1.Button18Click(Sender: TObject);
var
pinging : integer;
pinging_str : string;
begin
pinging := RadioGroup2.ItemIndex;
case pinging of
//
0:  pinging_str := '192.168.1.1';
//
1:   pinging_str := '8.8.8.8';
//
2:   pinging_str := 'ya.ru';
//
3:   pinging_str := 'nii_rzdpm.dnmu.ru';
//
4:   pinging_str := 'www.netlux.org';

  Else
   pinging_str := '192.168.1.1';
  end;
   pinging_str := '/k tracert.exe -4 ' + pinging_str;

//ключ /c если надо завершить cmd
ShellExecute(0,'open','cmd.exe',PWideChar(pinging_str),
'C:\Windows\system32\',SW_SHOW);
end;

procedure TForm1.Button19Click(Sender: TObject);
begin
//ключ /c если надо завершить cmd
ShellExecute(0,'open','cmd.exe','/k ipconfig /allcompartments |more',
'C:\Windows\system32\',SW_SHOW);
end;

procedure TForm1.Button1Click(Sender: TObject);
begin
Form1.Close;
end;

procedure TForm1.Button20Click(Sender: TObject);
var
 chosenDirectory : string;
 vibor_dir       : string;
 Stroka_pr : string;
begin
if SelectDirectory('Выбери папку','Компьютер',chosenDirectory) then
   begin
     vibor_dir :=chosenDirectory +'\' ;
   end
else
  begin
   ShowMessage('Выбор каталога прервался');
   Exit;
  end;
Form1.ComboBox6.Text := vibor_dir;

Stroka_pr := '--started-from-file  "'+ vibor_dir;
//Stroka_pr := '--started-from-file --playlist-enqueue "'+ vibor_dir;
ShellExecute(0,'open','vlc',
  PWideChar(Stroka_pr),nil,SW_SHOWNOACTIVATE);
end;

procedure TForm1.Button21Click(Sender: TObject);
var
  I: integer;
  Stroka_pr : string;
begin
 OpenDialog1.Options := [ofAllowMultiSelect, ofFileMustExist];
 OpenDialog1.Filter := 'All files (*.*)|*.*';
 OpenDialog1.FilterIndex := 2; { start the dialog showing all files }

// OpenDialog1.InitialDir := 'Компьютер';
 OpenDialog1.InitialDir := 'D:\download\download_from_torrent\';

 Stroka_pr := '--started-from-file  --playlist-enqueue ';
 if OpenDialog1.Execute then
       for i := 0 to openDialog1.Files.Count-1 do
          begin
           Stroka_pr := Stroka_pr +'"'+ openDialog1.Files[i]+'"  ';
          end;
//Memo2.Clear;
for i := 0 to openDialog1.Files.Count-1 do
  begin
   // Form1.Memo2.Lines.Add(openDialog1.Files[i]);
  end;
ShellExecute(0,'open','vlc',
  PWideChar(Stroka_pr),nil,SW_SHOWNOACTIVATE);
end;

procedure TForm1.Button22Click(Sender: TObject);
begin

Form5.Show;
//Form1.Hide;

end;

procedure TForm1.Button23Click(Sender: TObject);
begin
Form6.Show;
//Form1.Hide;
end;

procedure TForm1.Button24Click(Sender: TObject);
var
 i: Integer;
// sum : Integer;
// sel : byte;
  Stage      : integer;
  From_Date  : Tdate;
  To_Date    : Tdate;
begin
//
Stage := 0;
if DBGrid6.SelectedRows.Count > 0 then
 begin
// sum := 0;

 with DBGrid6.DataSource.DataSet do
 begin
 for i := 0 to DBGrid6.SelectedRows.Count-1 do
  begin
  GotoBookmark(DBGrid6.SelectedRows.Items[i]);
  From_Date  := DataModule1.FDQuery_Stage.FieldByName('From_Date').AsDateTime;
  To_Date    := DataModule1.FDQuery_Stage.FieldByName('To_Date').AsDateTime;
  Stage := Stage + DaysBetween(To_Date, From_Date  ) ;
  Form1.Edit11.Text := IntToStr(Stage);
  Form1.Edit13.Text := FormatAge(Now - StrToInt(Form1.Edit11.Text), Now, True);// 27 л. 7 мес. 24 д.
  end;
 end;


end;

end;

procedure TForm1.Button25Click(Sender: TObject);
begin
ShellExecute(Application.Handle, 'open', '"C:\Windows\System32\notepad.exe"', nil, nil,SW_SHOWNOACTIVATE);
end;

procedure TForm1.Button26Click(Sender: TObject);
var
// Text_Sql : string;
 chislo : string;
begin
chislo := Form1.Edit12.Text ;
Form1.Edit12.Text := FormatAge(Now - StrToInt(chislo), Now, True);// 27 л. 7 мес. 24 д.

end;

procedure TForm1.Button27Click(Sender: TObject);
begin
//
Form4.Show;
//Form1.Hide;
end;

procedure TForm1.Button28Click(Sender: TObject);
var
  Text_sql : string;
  {
  I: Integer;
   }
  nach_data :string;
  kon_data :string;

//  RecordSet   :integer;

begin
Text_sql := 'SELECT * FROM Bill_New  ';
Text_sql :=Text_sql +'  where Data BETWEEN :p1 and :p2 ';

nach_data :=FormatDateTime('yyyy-mm-dd', DateTimePicker1.Date);
kon_data  := FormatDateTime('yyyy-mm-dd', DateTimePicker2.Date);

DataModule1.FDQuery_excel.Active := false;
DataModule1.FDQuery_excel.SQL.Clear;
DataModule1.FDQuery_excel.SQL.Add(Text_SQL);

DataModule1.FDQuery_excel.Params.ParamByName('p1').Value :=nach_data  ;
DataModule1.FDQuery_excel.Params.ParamByName('p2').Value := kon_data;
DataModule1.FDQuery_excel.Active := true;

   Form1.DBGrid3.Columns[0].Width := 30;
   Form1.DBGrid3.Columns[0].Alignment:=taCenter;
   Form1.DBGrid3.Columns[1].Width := 80;
   Form1.DBGrid3.Columns[2].Width := 80;
   Form1.DBGrid3.Columns[3].Width := 80;
   Form1.DBGrid3.Columns[3].Alignment:=taCenter;
   Form1.DBGrid3.Columns[4].Width := 50;
   Form1.DBGrid3.Columns[4].Alignment:=taCenter;
   Form1.DBGrid3.Columns[5].Width := 80;
   Form1.DBGrid3.Columns[6].Width := 80;
   Form1.DBGrid3.Columns[7].Width := 80;
   Form1.DBGrid3.Columns[8].Width := 80;
   Form1.DBGrid3.Columns[9].Width := 80;
   Form1.DBGrid3.Columns[10].Width := 80;

   DataModule1.FDQuery_excel.Last;
//   RecordSet := DataModule1.FDQuery_excel.RecordCount;
   DataModule1.FDQuery_excel.last;
   Form1.DBGrid3.SetFocus;


end;

procedure TForm1.Button29Click(Sender: TObject);
begin
//
ShellExecute(Application.Handle, 'open', 'C:\windows\system32\mspaint.exe', nil, nil,SW_SHOWNOACTIVATE);
end;

procedure TForm1.Button2Click(Sender: TObject);
var
  FExcel, FSheet: Variant;
  Text_sql : string;
  I: Integer;
  nach_data :string;
  kon_data :string;
  RecordSet   :integer;
begin
//В Excel
Text_sql := 'SELECT * FROM Bill_New  ';
Text_sql :=Text_sql +'  where Data BETWEEN :p1 and :p2 ';

nach_data :=FormatDateTime('yyyy-mm-dd', DateTimePicker1.Date);
kon_data  := FormatDateTime('yyyy-mm-dd', DateTimePicker2.Date);

DataModule1.FDQuery_excel.Active := false;
DataModule1.FDQuery_excel.SQL.Clear;
DataModule1.FDQuery_excel.SQL.Add(Text_SQL);

DataModule1.FDQuery_excel.Params.ParamByName('p1').Value :=nach_data  ;
DataModule1.FDQuery_excel.Params.ParamByName('p2').Value := kon_data;
DataModule1.FDQuery_excel.Active := true;

   Form1.DBGrid3.Columns[0].Width := 30;
   Form1.DBGrid3.Columns[0].Alignment:=taCenter;
   Form1.DBGrid3.Columns[1].Width := 80;
   Form1.DBGrid3.Columns[2].Width := 80;
   Form1.DBGrid3.Columns[3].Width := 80;
   Form1.DBGrid3.Columns[3].Alignment:=taCenter;
   Form1.DBGrid3.Columns[4].Width := 50;
   Form1.DBGrid3.Columns[4].Alignment:=taCenter;
   Form1.DBGrid3.Columns[5].Width := 80;
   Form1.DBGrid3.Columns[6].Width := 80;
   Form1.DBGrid3.Columns[7].Width := 80;
   Form1.DBGrid3.Columns[8].Width := 80;
   Form1.DBGrid3.Columns[9].Width := 80;
   Form1.DBGrid3.Columns[10].Width := 80;

//Выводим в Excel
 if CheckExcelInstall =true then
  begin
    FExcel := CreateOleObject('Excel.Application');
    FExcel.DisplayAlerts := false;
    FExcel.Visible := true;
    FExcel.Application.EnableEvents := false;

    FExcel.WorkBooks.Add;
    FSheet := FExcel.Workbooks[1].WorkSheets[1];
    FSheet.Name := DateToStr(DateTimePicker1.Date)+' - '+DateToStr(DateTimePicker2.Date);
// Формируем заголовок Excel
//id Sobitie          DenNed     Data      Ostatok
//Pil   Travm  Gips  Zapl_inet  Sn_Gips  Den_Gizni|
   FSheet.Cells[1,1] := 'iD';

   FSheet.Cells[1,1].Font.Color := clWhite;
   FSheet.Cells[1,1].Font.Name := 'Times New Roman';
   FSheet.Cells[1,1].Font.Size := 14;
   FSheet.Cells[1,1].Interior.Color := rgb(0, 0, 0); // Первую строку закрашиваем в зелёный цвет, используя RGB
   FSheet.Cells[1,1].Font.Bold := True;

   FSheet.Cells[1,2] := 'Sob.';

   FSheet.Cells[1,2].Font.Color := clWhite;
   FSheet.Cells[1,2].Font.Name := 'Times New Roman';
   FSheet.Cells[1,2].Font.Size := 14;
   FSheet.Cells[1,2].Interior.Color := rgb(0, 0, 0); // Первую строку закрашиваем в зелёный цвет, используя RGB
   FSheet.Cells[1,2].Font.Bold := True;

   FSheet.Cells[1,3] := 'DN';

   FSheet.Cells[1,3].Font.Color := clWhite;
   FSheet.Cells[1,3].Font.Name := 'Times New Roman';
   FSheet.Cells[1,3].Font.Size := 14;
   FSheet.Cells[1,3].Interior.Color := rgb(0, 0, 0); // Первую строку закрашиваем в зелёный цвет, используя RGB
   FSheet.Cells[1,3].Font.Bold := True;

   FSheet.Cells[1,4] := 'Data';

   FSheet.Cells[1,4].Font.Color := clWhite;
   FSheet.Cells[1,4].Font.Name := 'Times New Roman';
   FSheet.Cells[1,4].Font.Size := 14;
   FSheet.Cells[1,4].Interior.Color := rgb(0, 0, 0); // Первую строку закрашиваем в зелёный цвет, используя RGB
   FSheet.Cells[1,4].Font.Bold := True;

   FSheet.Cells[1,5] := 'Ost.';

   FSheet.Cells[1,5].Font.Color := clWhite;
   FSheet.Cells[1,5].Font.Name := 'Times New Roman';
   FSheet.Cells[1,5].Font.Size := 14;
   FSheet.Cells[1,5].Interior.Color := rgb(0, 0, 0); // Первую строку закрашиваем в зелёный цвет, используя RGB
   FSheet.Cells[1,5].Font.Bold := True;

   FSheet.Cells[1,6] := 'Pil';

   FSheet.Cells[1,6].Font.Color := clWhite;
   FSheet.Cells[1,6].Font.Name := 'Times New Roman';
   FSheet.Cells[1,6].Font.Size := 14;
   FSheet.Cells[1,6].Interior.Color := rgb(0, 0, 0); // Первую строку закрашиваем в зелёный цвет, используя RGB
   FSheet.Cells[1,6].Font.Bold := True;

   FSheet.Cells[1,7] := 'Travm';

   FSheet.Cells[1,7].Font.Color := clWhite;
   FSheet.Cells[1,7].Font.Name := 'Times New Roman';
   FSheet.Cells[1,7].Font.Size := 14;
   FSheet.Cells[1,7].Interior.Color := rgb(0, 0, 0); // Первую строку закрашиваем в зелёный цвет, используя RGB
   FSheet.Cells[1,7].Font.Bold := True;

   FSheet.Cells[1,8] := 'Gips';

   FSheet.Cells[1,8].Font.Color := clWhite;
   FSheet.Cells[1,8].Font.Name := 'Times New Roman';
   FSheet.Cells[1,8].Font.Size := 14;
   FSheet.Cells[1,8].Interior.Color := rgb(0, 0, 0); // Первую строку закрашиваем в зелёный цвет, используя RGB
   FSheet.Cells[1,8].Font.Bold := True;

   FSheet.Cells[1,9] := 'ZIt';

   FSheet.Cells[1,9].Font.Color := clWhite;
   FSheet.Cells[1,9].Font.Name := 'Times New Roman';
   FSheet.Cells[1,9].Font.Size := 14;
   FSheet.Cells[1,9].Interior.Color := rgb(0, 0, 0); // Первую строку закрашиваем в зелёный цвет, используя RGB
   FSheet.Cells[1,9].Font.Bold := True;

   FSheet.Cells[1,10] := 'SG';

   FSheet.Cells[1,10].Font.Color := clWhite;
   FSheet.Cells[1,10].Font.Name := 'Times New Roman';
   FSheet.Cells[1,10].Font.Size := 14;
   FSheet.Cells[1,10].Interior.Color := rgb(0, 0, 0); // Первую строку закрашиваем в зелёный цвет, используя RGB
   FSheet.Cells[1,10].Font.Bold := True;

   FSheet.Cells[1,11] := 'DG';

   FSheet.Cells[1,11].Font.Color := clWhite;
   FSheet.Cells[1,11].Font.Name := 'Times New Roman';
   FSheet.Cells[1,11].Font.Size := 14;
   FSheet.Cells[1,11].Interior.Color := rgb(0, 0, 0); // Первую строку закрашиваем в зелёный цвет, используя RGB
   FSheet.Cells[1,11].Font.Bold := True;

   FSheet.Cells[1,12] := 'Zam';

   FSheet.Cells[1,12].Font.Color := clWhite;
   FSheet.Cells[1,12].Font.Name := 'Times New Roman';
   FSheet.Cells[1,12].Font.Size := 14;
   FSheet.Cells[1,12].Interior.Color := rgb(0, 0, 0); // Первую строку закрашиваем в зелёный цвет, используя RGB
   FSheet.Cells[1,12].Font.Bold := True;


   FSheet.Columns[1].ColumnWidth := 3;
   FSheet.Columns[2].ColumnWidth := 15;
   FSheet.Columns[3].ColumnWidth := 15;
   FSheet.Columns[4].ColumnWidth := 10;

   FSheet.Columns[6].ColumnWidth := 12;

   FSheet.Columns[12].ColumnWidth := 30;


  // FSheet.Rows[1].Font.Name := 'Times New Roman';
  // FSheet.Rows[1].Font.Size := 14;
  // FSheet.Rows[1].Font.Bold := True;
  // FSheet.Rows[1].Font.Color := clWhite;
  FSheet.Rows[1].HorizontalAlignment := 3;
  // FSheet.Rows[1].Interior.Color := rgb(0, 0, 0); // Первую строку закрашиваем в зелёный цвет, используя RGB

   DataModule1.FDQuery_excel.Last;
   RecordSet := DataModule1.FDQuery_excel.RecordCount;
   DataModule1.FDQuery_excel.First;

  for I := 2 to RecordSet +1 do
     begin
      FSheet.Cells[I,1] := DataModule1.FDQuery_excel.FieldByName('id').AsString;
      FSheet.Cells[I,1].Interior.Color := rgb(192, 255, 192);

      FSheet.Cells[I,2] := DataModule1.FDQuery_excel.FieldByName('Sobitie').AsString;
      if (DataModule1.FDQuery_excel.FieldByName('Sobitie').AsString= 'Псрл')
         then   FSheet.Cells[I,2].Interior.Color := RGB(0, 176, 240);

      FSheet.Cells[I,3] := DataModule1.FDQuery_excel.FieldByName('DenNed').AsString;
         FSheet.Cells[I,3].Interior.Color := RGB(244, 176, 132);
         FSheet.Cells[I,3].HorizontalAlignment := 4;

      FSheet.Cells[I,4] := DataModule1.FDQuery_excel.FieldByName('Data').AsString;
         FSheet.Cells[I,4].HorizontalAlignment := 4;

      FSheet.Cells[I,5] := DataModule1.FDQuery_excel.FieldByName('Ostatok').AsString;
         FSheet.Cells[I,5].HorizontalAlignment := 3;

      FSheet.Cells[I,6] := DataModule1.FDQuery_excel.FieldByName('Pil').AsString;
          if (DataModule1.FDQuery_excel.FieldByName('Pil').AsString = 'НЕ ПИЛ')
             then   FSheet.Cells[I,6].Interior.Color := RGB(50, 193, 47);
          if (DataModule1.FDQuery_excel.FieldByName('Pil').AsString = 'ПИЛ')
             then   FSheet.Cells[I,6].Interior.Color := RGB(255, 0, 0);
          if (DataModule1.FDQuery_excel.FieldByName('Pil').AsString = 'По разному')
             then   FSheet.Cells[I,6].Interior.Color := RGB(255, 255, 0);

             FSheet.Cells[I,6].HorizontalAlignment := 3;
      FSheet.Cells[I,7] := DataModule1.FDQuery_excel.FieldByName('Travm').AsString;
         FSheet.Cells[I,7].Interior.Color := RGB(189, 215, 238);
         FSheet.Cells[I,7].HorizontalAlignment := 3;
      FSheet.Cells[I,8] := DataModule1.FDQuery_excel.FieldByName('Gips').AsString;
         FSheet.Cells[I,8].Interior.Color := RGB(155, 194, 230);
         FSheet.Cells[I,8].HorizontalAlignment := 3;
      FSheet.Cells[I,9] := DataModule1.FDQuery_excel.FieldByName('Zapl_inet').AsString;
         FSheet.Cells[I,9].Interior.Color := RGB(47, 117, 181);
         FSheet.Cells[I,9].HorizontalAlignment := 3;

      FSheet.Cells[I,10] := DataModule1.FDQuery_excel.FieldByName('Sn_Gips').AsString;
         FSheet.Cells[I,10].Interior.Color := RGB(31, 78, 120);
         FSheet.Cells[I,10].HorizontalAlignment := 3;
         FSheet.Cells[I,10].Font.Color := clWhite;

      FSheet.Cells[I,11] := DataModule1.FDQuery_excel.FieldByName('Den_Gizni').AsString;
         FSheet.Cells[I,11].HorizontalAlignment := 3;

      FSheet.Cells[I,12] := DataModule1.FDQuery_excel.FieldByName('Zam').AsString;
       if (DataModule1.FDQuery_excel.FieldByName('Zam').AsString <> '')
       then
         begin
        FSheet.Cells[I,12].Interior.Color := RGB(198, 0, 159);
        FSheet.Cells[I,12].Font.Bold := True;
        FSheet.Cells[I,12].Font.Size := 14;
         end;
       // Font.Size := 14;

      DataModule1.FDQuery_excel.Next;
     end;
 FExcel.WindowState := -4137;

  end

else
  begin
     MessageBox(0,'Приложение MS Excel не установлено на этом компьютере','Ошибка',MB_OK+MB_ICONERROR);
  end;
end;

procedure TForm1.Button30Click(Sender: TObject);
begin
//
 ShellExecute(Application.Handle, 'open', 'http://netlux.org/tv-gid/', nil, nil,SW_SHOWNOACTIVATE);
// ShellExecute(Application.Handle, 'open', 'http://lk.netlux.org/Account/Services', nil, nil,SW_SHOWNOACTIVATE);
end;

procedure TForm1.Button31Click(Sender: TObject);
label
  stroka;
var
Path       : String;
Sr         : TSearchRec;
F,F_n      : TextFile; //файловая переменная
S,S1,S3    : string;  //текстовая строка
SL         :TStringList;
Search,
SearchName: String;
Input,Leng, i: Integer;

fileName   : string;
fileDate   : Integer;
Path_1   :string;
Path_arch :string;
//для вытягивания файла
FileOnNet, LocalFileName: string;


begin

Path := 'D:\';
//Поиск файла и удаление
if FindFirst(Path+'unicast.m3u', faAnyFile, sr) = 0 then
  begin
  // Должен освободить ресурсы, используемые этими успешными, поисками
    DeleteFile(Path+sr.Name);
  end;
  FindClose(sr);
//=========================================
Path_1 := ExtractFileDir(Application.ExeName)+'\EGP_cur\';
     if System.SysUtils.DirectoryExists(Path_1) then
       begin
         path_arch := Path_1;
       end
     else
       begin
         CreateDir(Path_1);
         path_arch := Path_1;
       end;

Path := 'D:\download\';
//Поиск файла и удаление
if FindFirst(Path+'unicast.m3u', faAnyFile, sr) = 0 then
  begin
  // Должен освободить ресурсы, используемые этими успешными, поисками
    DeleteFile(Path+sr.Name);
  end;
  FindClose(sr);
//=========================================
//===========================================================
  FileOnNet := 'http://iptv.netlux.org/unicast.m3u';
 // if DirectoryExists('D:\md') then DeleteDir('D:\md');
//Создаем заново

Path_1 := ExtractFileDir(Application.ExeName)+'\EGP_cur\';
     if System.SysUtils.DirectoryExists(Path_1) then
       begin
         path_arch := Path_1;
       end
     else
       begin
         CreateDir(Path_1);
         path_arch := Path_1;
       end;


  LocalFileName := Path_1+ 'unicast.m3u';

  if GetInetFile(FileOnNet, LocalFileName) = true then

    begin
    Form1.Edit14.Color := clHotLight ;
    //ShowMessage('Download unicast.m3u successful')
    end
  else
    ShowMessage('Error load file unicast.m3u download');
//===========================================================

//ShellExecute(Application.Handle, 'open', 'http://iptv.netlux.org/unicast.m3u', nil, nil,SW_SHOWNOACTIVATE);
//Sleep(5000);

//редактируем unicast.m3u
SL:= TStringList.Create;
SL.LoadFromFile(Path_1+ 'unicast.m3u');
//SL.Sorted := true;
//"Россия 1"
//5 канал
if Form1.CheckBox1.Checked = true then
  begin
   Form1.CheckBox2.Enabled := false;
   Form1.CheckBox3.Enabled := false;
   Search:='zvezda';
  end;
if Form1.CheckBox2.Checked = true then
  begin
   Form1.CheckBox1.Enabled := false;
   Form1.CheckBox3.Enabled := false;
   Search:='rus1';
  end;

if Form1.CheckBox3.Checked = true then
  begin
   Form1.CheckBox1.Enabled := false;
   Form1.CheckBox2.Enabled := false;
   Search:='5ch';
  end;

if (Form1.CheckBox1.Checked <> true) and
   (Form1.CheckBox2.Checked <> true) and
   (Form1.CheckBox3.Checked <> true) then Search:='zvezda';


//Leng:=Length(Search);
for i := 0 to SL.Count - 1 do                      //Выполняем пока нет конца файла
begin
  Input:=Pos(Search, SL.Strings[i]);                       //Поиск вхождения слова
  if Input>0 then
    begin
     S:=SL.Strings[i];
     S1 := SL.Strings[i+1];
     Goto stroka;
    end;
end;

Stroka:
//Редактируем файл
AssignFile(F,Path_1+ 'unicast.m3u'); //связывает F с именем файла
AssignFile(F_n,'D:\unicast.m3u'); //связывает F_n с именем файла
Reset(F); //открывает существующий файл
ReWrite(F_n); //создает новый файл и открывает его
i := 0;
while (not EOF(f)) do begin
  Readln(f, s3);
  if i = 1 then
  begin
  Writeln(f_n,S);
  Writeln(f_n,S1);
  end;
  Writeln(f_n,S3);
  Inc(i);
end;
CloseFile(F);
CloseFile(F_n);

//CopyFile('D:\unicast.m3u', 'D:\New_Bill_Tab\EPG_NetLux\EGP_Cur\unicast.m3u', false);

  fileName := 'D:\unicast.m3u';
  fileDate := FileAge(fileName);
   if fileDate > -1 then
   Form1.Edit14.Text :=( 'посл.изм   ' +
     DateToStr(FileDateToDateTime(fileDate)));

  if DaysBetween(FileDateToDateTime(fileDate),now) > 1 then
    begin
     Form1.Edit14.Color := clRed ;
    end;

end;

procedure TForm1.Button3Click(Sender: TObject);
begin
 ShellExecute(Application.Handle, 'open', 'http://lk.netlux.org/Account/Services', nil, nil,SW_SHOWNOACTIVATE);
end;

procedure TForm1.Button40Click(Sender: TObject);
begin
ShellExecute(Application.Handle, 'open', 'C:\Program Files\VideoLAN\VLC\vlc.exe',
  'D:\Download\DownLoad_From_Torrent\Сегодня', nil,SW_SHOWNOACTIVATE);
end;

procedure TForm1.Button41Click(Sender: TObject);
begin
 ClearDir( 'D:\Download\DownLoad_From_Torrent\Сегодня' );
 if DirectoryExists('d:\md') then
  begin
   if ClearDir( 'd:\md' )  then ShowMessage('d:\md очищен')

   else ShowMessage('d:\md не удален, произошла ошибка : '+IntToStr(GetLastError));

  end;
end;

procedure TForm1.Button4Click(Sender: TObject);
begin
//
ShellExecute(Application.Handle, 'open', '"C:\Program Files\Notepad++\notepad++.exe"', nil, nil,SW_SHOWNOACTIVATE);
end;

procedure TForm1.Button5Click(Sender: TObject);
begin
  Form3.Show;
  Form1.Hide;

end;

procedure TForm1.Button6Click(Sender: TObject);
var
  i :integer;
  url       : string;
  page      : string;
  find      : string;
  position  : integer;
  position_last  : integer;
  copy_pos       : integer;
  S2 : String;   //строка для удаления
  SL:TStringList;
  s1:string;
begin
// Гороскоп
//url:='https://orakul.com/horoscope/astrologic/general/gemini/today.html ';
url:='https://horo.mail.ru/amp/prediction/gemini/today/ ';
//https://horo.mail.ru/amp/prediction/gemini/today/
{
page:=IdHTTP1.get(url);  //скачиваем код страницы
position := AnsiPos('Общий гороскоп на сегодня,', page);
position_last := AnsiPos('Подробнее', page);
copy_pos := position_last - position;
find := Copy(page, position, copy_pos);
//чистим страницу
S2 :=  '</h2>';
find := StringReplace(find, S2,'', [rfReplaceAll, rfIgnoreCase]);

S2 :=  '</div>';
find := StringReplace(find, S2,'', [rfReplaceAll, rfIgnoreCase]);

S2 :=  '<div class="horoBlock">';
find := StringReplace(find, S2,'', [rfReplaceAll, rfIgnoreCase]);

S2 :=  '<p class="">';
find := StringReplace(find, S2,'', [rfReplaceAll, rfIgnoreCase]);

S2 :=  '<p id="morespeciallink"><a href=';
find := StringReplace(find, S2,'', [rfReplaceAll, rfIgnoreCase]);

S2 :=  '</p>';
find := StringReplace(find, S2,'', [rfReplaceAll, rfIgnoreCase]);

S2 :=  '>';
find := StringReplace(find, S2,'', [rfReplaceAll, rfIgnoreCase]);

S2 :=  '"https://orakul.com/horoscope/astrologic/more/gemini/today.html"';
find := StringReplace(find, S2,'', [rfReplaceAll, rfIgnoreCase]);

SL := TStringList.Create;

//ExtractStrings([#10, #13], [' '],  find , SL);
  try
  SL.Delimiter := #13;
  SL.StrictDelimiter := True;
  SL.DelimitedText := find;
  Memo1.Clear;
for i:=SL.Count-1 downto 0
    do if Length(trim(SL.Strings[i]))=0
       then SL.Delete(i);

  for i := 0 to sl.Count - 1 do
    begin
       s1  := Trim(SL[i]);
    Memo1.Lines.Add(WrapText(s1,90));
    end;
    //end;

  finally
    SL.Free
  end;
    // Memo1.Lines.Assign(SL);
//  for i := 0 to sl.Count - 1 do
}
  ShellExecute(Application.Handle, 'open','https://horo.mail.ru/prediction/gemini/today/ ', nil, nil,SW_SHOWNOACTIVATE);
//Form1.Button38.Visible := True;
end;

procedure TForm1.Button7Click(Sender: TObject);
var
Text_SQL : String;
id       : integer;
DenNed   : string;
day_nedeli    : array[1..7] of string;
begin
 day_nedeli[1] := 'Понедельник';
 day_nedeli[2] := 'Вторник';
 day_nedeli[3] := 'Среда';
 day_nedeli[4] := 'Четверг';
 day_nedeli[5] := 'Пятница';
 day_nedeli[6] := 'Суббота';
 day_nedeli[7] := 'Воскресенье';

DenNed:= day_nedeli[DayOfTheWeek( StrToDate(today_str))] ;
//=======================================================
//Вычислем id сегодняшней записи

Text_SQL := 'select * from  Bill_New  where "Data" = :p100';
Datamodule1.FDQuery_ins.Active :=false;
Datamodule1.FDQuery_ins.SQL.Clear;
Datamodule1.FDQuery_ins.SQL.Add(Text_SQL);

Datamodule1.FDQuery_ins.Params.ParamByName('p100').Value := DateToBase(today_str);
Datamodule1.FDQuery_ins.Active :=true;
id:=Datamodule1.FDQuery_ins.FieldByName('id').AsInteger;
//========================================================
//Update base
Text_SQL := 'UPDATE Bill_New SET ';
Text_SQL :=Text_SQL + 'Ostatok = :p1,';
Text_SQL :=Text_SQL + 'Pil=:p3,';
Text_SQL :=Text_SQL +'Travm = :p4,';
Text_SQL :=Text_SQL +'Gips = :p5,';
Text_SQL :=Text_SQL +'Sn_Gips = :p6,';
Text_SQL :=Text_SQL +'Zapl_inet = :p7,';
Text_SQL :=Text_SQL +'Sobitie = :p8,';
Text_SQL :=Text_SQL +'Den_Gizni = :p9,';
Text_SQL :=Text_SQL +'DenNed   = :p10,';
Text_SQL :=Text_SQL +'Data   = :p11';
Text_SQL :=Text_SQL + '  WHERE id = :p2;';

Datamodule1.FDQuery_ins.Active :=false;
Datamodule1.FDQuery_ins.SQL.Clear;
Datamodule1.FDQuery_ins.SQL.Add(Text_SQL);
Datamodule1.FDQuery_ins.Params.ParamByName('p1').Value := IntToStr(Form1.SpinEdit3.Value);
Datamodule1.FDQuery_ins.Params.ParamByName('p2').Value := IntToStr(id) ;
Datamodule1.FDQuery_ins.Params.ParamByName('p3').Value := Form1.ComboBox1.Text;
Datamodule1.FDQuery_ins.Params.ParamByName('p4').Value := Form1.Edit2.Text;
Datamodule1.FDQuery_ins.Params.ParamByName('p5').Value := Form1.Edit3.Text;
Datamodule1.FDQuery_ins.Params.ParamByName('p6').Value := Form1.Edit4.Text;
//!!!!!!!!
Datamodule1.FDQuery_ins.Params.ParamByName('p7').Value :=
 Copy(Form1.Edit5.Text,1,3);

Datamodule1.FDQuery_ins.Params.ParamByName('p8').Value := Form1.ComboBox2.Text;
Datamodule1.FDQuery_ins.Params.ParamByName('p9').Value := Form1.Edit7.Text
+' или  '+Form1.Edit8.Text;
Datamodule1.FDQuery_ins.Params.ParamByName('p10').Value := DenNed;
Datamodule1.FDQuery_ins.Params.ParamByName('p11').Value := DateToBase(today_str);
Datamodule1.FDQuery_ins.ExecSQL;

   Text_SQL :='select * from Bill_New ';

Datamodule1.FDQuery_osn.Active :=false;
Datamodule1.FDQuery_osn.SQL.Clear;
Datamodule1.FDQuery_osn.SQL.Add(Text_SQL);
Datamodule1.FDQuery_osn.Active :=true;

   Form1.DBGrid2.Columns[0].Width := 50;
   Form1.DBGrid2.Columns[0].Alignment:=taCenter;
   Form1.DBGrid2.Columns[1].Width := 50;
   Form1.DBGrid2.Columns[1].Alignment:=taCenter;
   Form1.DBGrid2.Columns[2].Width := 105;
   Form1.DBGrid2.Columns[3].Width := 80;
   Form1.DBGrid2.Columns[3].Alignment:=taCenter;
   Form1.DBGrid2.Columns[4].Width := 50;
   Form1.DBGrid2.Columns[4].Alignment:=taCenter;
//   Form1.DBGrid2.Columns[4].Font
   Form1.DBGrid2.Columns[5].Width := 80;
   Form1.DBGrid2.Columns[6].Width := 80;
   Form1.DBGrid2.Columns[7].Width := 80;
   Form1.DBGrid2.Columns[8].Width := 80;
   Form1.DBGrid2.Columns[9].Width := 80;
   Form1.DBGrid2.Columns[10].Width := 210;

   Form1.Edit1.Text := today_str;
   Edit1Change(Sender);

   Form1.DBGrid2.SetFocus;

end;

procedure TForm1.Button8Click(Sender: TObject);
begin
//Form1.Hide;
Form11.Show;
end;

procedure TForm1.Button9Click(Sender: TObject);
begin
ShellExecute(0,'open','Calc.exe','','C:\Windows\system32\',SW_SHOW);

end;

procedure TForm1.Button_DengiClick(Sender: TObject);
var
Summa_p         :string;
Nomer_p         :string;
Data_Op_p       :string;
Data_Op_p_dt    :TDateTime;
Detali_p        :string;
Ostatok_Zarpl_p :string;
Contragent_p    :string;
//-----------------------
MyExcel, MySheet: Variant;
filename : string;
row,col,r,c  :integer;
Text_sql    : string;

begin
if PageControl2.ActivePage = TabSheet12 then  //(Zarpl)
 begin
   OpenDialog1.InitialDir := GetCurrentDir;
   OpenDialog1.Filter := 'Excel файлы|z*.x*|Все файлы|*.*';
   if Form1.OpenDialog1.Execute then
   begin
     if CheckExcelInstall =true then
     begin
     filename := OpenDialog1.FileName;
     end;
       MyExcel := CreateOleObject('Excel.Application');
       MyExcel.DisplayAlerts := false;
       MyExcel.Visible := false;
       MyExcel.Application.EnableEvents := false;
       MyExcel.WindowState := -4137;
       MyExcel.Workbooks.Open(filename);

       MySheet := MyExcel.Workbooks[1].WorkSheets[1];
  //Последняя ипользованная ячейка
       row := MySheet.UsedRange.Rows.Count;
       col :=  MySheet.UsedRange.Columns.Count;

for r := row DownTo 2 do
     begin
      Summa_p:=MySheet.Cells[r,1].Text;
      Summa_p:= StringReplace(Summa_p, ' ', '', [rfReplaceAll]);
      Summa_p:= StringReplace(Summa_p, ',', '.', [rfReplaceAll]);

      Data_Op_p := MySheet.Cells[r,2].Text;
      Data_Op_p_dt := StrToDate(Data_Op_p);
      Data_Op_p := FormatDateTime('yyyy-mm-dd',Data_Op_p_dt);

      Nomer_p := MySheet.Cells[r,3].Text;

      Detali_p:=MySheet.Cells[r,4].Text;

      Text_sql :=
   'insert into Vipiska_Zarpl  (Summa,Nomer,Data_Op,Detali) values(:p1,:p2,:p3,:p4)';

      DataModule1.FDQuery_Den_Inp.Active := False;
      DataModule1.FDQuery_Den_Inp.SQL.Clear;
      DataModule1.FDQuery_Den_Inp.SQL.Add(Text_sql) ;
      DataModule1.FDQuery_Den_Inp.Params.ParamByName('p1').Value
        := Summa_p;
      DataModule1.FDQuery_Den_Inp.Params.ParamByName('p2').Value
        := Nomer_p;
      DataModule1.FDQuery_Den_Inp.Params.ParamByName('p3').Value
        := Data_Op_p;
      DataModule1.FDQuery_Den_Inp.Params.ParamByName('p4').Value
        := Detali_p;
//      DataModule1.FDQuery_Den_Inp.Params.ParamByName('p5').Value
//        := Contragent_p;
      DataModule1.FDQuery_Den_Inp.ExecSQL;
     end;


  //  FExcel.ActiveWorkbook.SaveCopyAs(path_sohran+'\M2D'+seg_data+'_'+seg_time+'.xlsx');
    // StopExcel;
    DataModule1.FDQuery_Dengi_Zarpl.Active := false;
    DataModule1.FDQuery_Dengi_Zarpl.Active := true;
    DataModule1.FDQuery_Dengi_Zarpl.Last;
    Form1.DBGrid10.SetFocus;



  MyExcel.Quit; //закрываем Excel
  MyExcel := Unassigned;//отсоединяемся от экземпляра
   end;
 end;

if PageControl2.ActivePage = TabSheet13 then  //(Pens)
 begin
   OpenDialog1.InitialDir := GetCurrentDir;
   OpenDialog1.Filter := 'Excel файлы|p*.x*|Все файлы|*.*';
   if Form1.OpenDialog1.Execute then
   begin
     if CheckExcelInstall =true then
     begin
     filename := OpenDialog1.FileName;
     end;
       MyExcel := CreateOleObject('Excel.Application');
       MyExcel.DisplayAlerts := false;
       MyExcel.Visible := false;
       MyExcel.Application.EnableEvents := false;
       MyExcel.WindowState := -4137;
       MyExcel.Workbooks.Open(filename);

       MySheet := MyExcel.Workbooks[1].WorkSheets[1];
  //Последняя ипользованная ячейка
       row := MySheet.UsedRange.Rows.Count;
       col :=  MySheet.UsedRange.Columns.Count;

for r := row DownTo 2 do
     begin
      Summa_p:=MySheet.Cells[r,1].Text;
      Summa_p:= StringReplace(Summa_p, ' ', '', [rfReplaceAll]);
      Summa_p:= StringReplace(Summa_p, ',', '.', [rfReplaceAll]);

      Data_Op_p := MySheet.Cells[r,2].Text;
      Data_Op_p_dt := StrToDate(Data_Op_p);
      Data_Op_p := FormatDateTime('yyyy-mm-dd',Data_Op_p_dt);

      Nomer_p := MySheet.Cells[r,3].Text;

      Detali_p:=MySheet.Cells[r,4].Text;

      Text_sql :=
   'insert into Vipiska_Pens  (Summa,Nomer,Data_Op,Detali) values(:p1,:p2,:p3,:p4)';

      DataModule1.FDQuery_Den_Inp.Active := False;
      DataModule1.FDQuery_Den_Inp.SQL.Clear;
      DataModule1.FDQuery_Den_Inp.SQL.Add(Text_sql) ;
      DataModule1.FDQuery_Den_Inp.Params.ParamByName('p1').Value
        := Summa_p;
      DataModule1.FDQuery_Den_Inp.Params.ParamByName('p2').Value
        := Nomer_p;
      DataModule1.FDQuery_Den_Inp.Params.ParamByName('p3').Value
        := Data_Op_p;
      DataModule1.FDQuery_Den_Inp.Params.ParamByName('p4').Value
        := Detali_p;
//      DataModule1.FDQuery_Den_Inp.Params.ParamByName('p5').Value
//        := Contragent_p;
      DataModule1.FDQuery_Den_Inp.ExecSQL;



     end;


  //  FExcel.ActiveWorkbook.SaveCopyAs(path_sohran+'\M2D'+seg_data+'_'+seg_time+'.xlsx');
    // StopExcel;
    DataModule1.FDQuery_Dengi_Pens.Active := false;
    DataModule1.FDQuery_Dengi_Pens.Active := true;
    DataModule1.FDQuery_Dengi_Pens.Last;
    Form1.DBGrid11.SetFocus;
  //  FExcel.ActiveWorkbook.SaveCopyAs(path_sohran+'\M2D'+seg_data+'_'+seg_time+'.xlsx');
  //   StopExcel;
    MyExcel.Quit; //закрываем Excel
    MyExcel := Unassigned;//отсоединяемся от экземпляра


   end;

  end;
 end;

procedure TForm1.Button_OchClick(Sender: TObject);
var
  Text_sql :  string;
begin
//Zarpl
if (Form1.PageControl2.ActivePage =TabSheet12) and
   (Form1.Button_Dengi.Caption = 'Загрузить выписку по зарплате') then
 begin
   Text_sql :='delete FROM Vipiska_Zarpl';
   DataModule1.FDQuery_Den_Inp.Active := false;
   DataModule1.FDQuery_Den_Inp.SQL.Clear;
   DataModule1.FDQuery_Den_Inp.SQL.Add(Text_sql);
   DataModule1.FDQuery_Den_Inp.ExecSQL;
   Text_sql :='delete  FROM sqlite_sequence WHERE name = :p1';
   DataModule1.FDQuery_Den_Inp.Active := false;
   DataModule1.FDQuery_Den_Inp.SQL.Clear;
   DataModule1.FDQuery_Den_Inp.SQL.Add(Text_sql);
   DataModule1.FDQuery_Den_Inp.Params.ParamByName('p1').Value  := 'Vipiska_Zarpl';
   DataModule1.FDQuery_Den_Inp.ExecSQL;
   DataModule1.FDQuery_Dengi_Zarpl.Active := false;
   DataModule1.FDQuery_Dengi_Zarpl.Active := true;


 end;
//Pens
if (Form1.PageControl2.ActivePage =TabSheet13) and
   (Form1.Button_Dengi.Caption = 'Загрузить выписку по пенсии') then
 begin
   Text_sql :='delete FROM Vipiska_Pens';
   DataModule1.FDQuery_Den_Inp.Active := false;
   DataModule1.FDQuery_Den_Inp.SQL.Clear;
   DataModule1.FDQuery_Den_Inp.SQL.Add(Text_sql);
   DataModule1.FDQuery_Den_Inp.ExecSQL;
   Text_sql :='delete  FROM sqlite_sequence WHERE name = :p1';
   DataModule1.FDQuery_Den_Inp.Active := false;
   DataModule1.FDQuery_Den_Inp.SQL.Clear;
   DataModule1.FDQuery_Den_Inp.SQL.Add(Text_sql);
   DataModule1.FDQuery_Den_Inp.Params.ParamByName('p1').Value  := 'Vipiska_Pens';
   DataModule1.FDQuery_Den_Inp.ExecSQL;
   DataModule1.FDQuery_Dengi_Pens.Active := false;
   DataModule1.FDQuery_Dengi_Pens.Active := true;

   end;

end;

procedure TForm1.Button_oplClick(Sender: TObject);
begin
  Form2.show;
//  Form1.Hide;
end;

procedure TForm1.ClipBoard1Click(Sender: TObject);
begin
//DBGrid1  - оплата интернета
if PopupMenu1.PopupComponent = DBGrid1 then
  CopySelGRDToClb(DBGrid1);
//DBGrid2  - основная таблица
if PopupMenu1.PopupComponent = DBGrid2 then
  CopySelGRDToClb(DBGrid2);
//DBGrid2  - редактирование основной таблицы
if PopupMenu1.PopupComponent = Form3.DBGrid2 then
   CopySelGRDToClb(Form3.DBGrid2);
//DBGrid3  - таилица выборки
if PopupMenu1.PopupComponent = Form1.DBGrid3 then
   CopySelGRDToClb(Form1.DBGrid3);
//DBGrid4 - таблица электроэнергии
if PopupMenu1.PopupComponent = DBGrid4 then
  CopySelGRDToClb(DBGrid4);
//DBGrid5  - таблица оплаты коммнулки
if PopupMenu1.PopupComponent = DBGrid5 then
  CopySelGRDToClb(DBGrid5);
//DBGrid6 - таблица моего стажа
if PopupMenu1.PopupComponent = DBGrid6 then
  CopySelGRDToClb(DBGrid6);
//DBGrid7 - таблица паролей
if PopupMenu1.PopupComponent = DBGrid7 then
  CopySelGRDToClb(DBGrid7);

end;

procedure TForm1.ComboBox10Select(Sender: TObject);
var
  S,chislo,mes,god: string;
  it       : integer;
begin
it := Form1.ComboBox10.ItemIndex;
case it of
//
0: begin
     chislo := Copy(DateToStr(Now),1,2);
     mes := Months[MonthOf(Now)];
     god := Copy(DateToStr(Now),7,4);
     S := 'мульткалендарь'+' '+chislo+' '+mes+' '+god ;
     //S := 'https://www.youtube.com/results?search_query=' + S;
     S := TNetEncoding.URL.Encode(S);
     S := 'https://www.youtube.com/results?search_query=' + S;
     Perehod(S);
   end;

//
1:  begin
     chislo := Copy(DateToStr(Now),1,2);
     mes := Months[MonthOf(Now)];
     god := Copy(DateToStr(Now),7,4);
     S := 'молитва'+' '+chislo+' '+mes+' '+god ;
     //S := 'https://www.youtube.com/results?search_query=' + S;
    // S := TNetEncoding.URL.Encode(S);
     S := 'https://rutube.ru/search/?suggest=1&query=' + S;
     Perehod(S);
   end;

//
2:  begin
     chislo := Copy(DateToStr(Now),1,2);
     mes := Months[MonthOf(Now)];
     god := Copy(DateToStr(Now),7,4);
     S := 'МОЛИТВА АНГЕЛУ ХРАНИТЕЛЮ'+' '+chislo+' '+mes+' '+god  ;
     //S := 'https://www.youtube.com/results?search_query=' + S;
     //S := TNetEncoding.URL.Encode(S);
     S := 'https://rutube.ru/search/?suggest=1&query=' + S;
     Perehod(S);
   end;
//
4 :  begin
     chislo := Copy(DateToStr(Now),1,2);
     mes := Months[MonthOf(Now)];
     god := Copy(DateToStr(Now),7,4);
     S := 'молитва'+' '+chislo+' '+mes+' '+god ;
     //S := 'https://www.youtube.com/results?search_query=' + S;
     S := TNetEncoding.URL.Encode(S);
     S := 'https://www.youtube.com/results?search_query=' + S;
     Perehod(S);
   end;
//
5:  begin
     chislo := Copy(DateToStr(Now),1,2);
     mes := Months[MonthOf(Now)];
     god := Copy(DateToStr(Now),7,4);
     S := 'Молитва Ангелу Хранителю'+' '+chislo+' '+mes+' '+god ;  ;
     //S := 'https://www.youtube.com/results?search_query=' + S;
     S := TNetEncoding.URL.Encode(S);
     S := 'https://www.youtube.com/results?search_query=' + S;
     Perehod(S);
   end;



//
//3:


  Else
    ShellExecute(Application.Handle, 'open', 'http://www.vsetv.com/', nil, nil,SW_SHOWNOACTIVATE);
  end;

end;

procedure TForm1.ComboBox11Select(Sender: TObject);
var
S,chislo,mes,god: string;
it       : integer;
begin
it := Form1.ComboBox11.ItemIndex;
  case it of
   // ПСБ Маркет
0: begin
     //chislo := Copy(DateToStr(Now),1,2);
     //mes := Months[MonthOf(Now)];
     //god := Copy(DateToStr(Now),7,4);
     //S := 'Самые шокирующие гипотезы'+' '+chislo+' '+mes+' '+god ;
     //S := 'https://www.youtube.com/results?search_query=' + S;
     //S := TNetEncoding.URL.Encode(S);
     S := 'https://xn--80abntiqkep.xn--p1ai/' + S;
     Perehod(S);
   end;
   // G-STORE
1: begin
     //chislo := Copy(DateToStr(Now),1,2);
     //mes := Months[MonthOf(Now)];
     //god := Copy(DateToStr(Now),7,4);
     //S := 'Загадки человечества с Олегом Шишкиным'+' '+chislo+' '+mes+' '+god ;
     //S := 'https://www.youtube.com/results?search_query=' + S;
     //S := TNetEncoding.URL.Encode(S);
     S := 'https://www.g-store.org/' + S;
     Perehod(S);
   end;

   // Техно Макс
2: begin
     //chislo := Copy(DateToStr(Now),1,2);
     //mes := Months[MonthOf(Now)];
     //god := Copy(DateToStr(Now),7,4);
     //S := 'Тайны Чапман'+' '+chislo+' '+mes+' '+god ;
     //S := 'https://www.youtube.com/results?search_query=' + S;
     //S := TNetEncoding.URL.Encode(S);
     S := 'https://tehnomaxplus.ru/' + S;
     Perehod(S);
   end;

   // uni-marketplace.com
3: begin
    // chislo := Copy(DateToStr(Now),1,2);
    // mes := Months[MonthOf(Now)];
    // god := Copy(DateToStr(Now),7,4);
    // S := '«Улика из прошлого»'+' '+chislo+' '+mes+' '+god ;
     //S := 'https://www.youtube.com/results?search_query=' + S;
    // S := TNetEncoding.URL.Encode(S);
     S := 'https://uni-marketplace.com/' + S;
     Perehod(S);
   end;

   // Ф-Март
4: begin
     //chislo := Copy(DateToStr(Now),1,2);
     //mes := Months[MonthOf(Now)];
     //god := Copy(DateToStr(Now),7,4);
     //S := 'Главные тайны мира'+' '+chislo+' '+mes+' '+god ;
     //S := 'https://www.youtube.com/results?search_query=' + S;
     //S := TNetEncoding.URL.Encode(S);
     S := 'https://fmart-shop.ru/' + S;
     Perehod(S);
   end;

   //  Фокс
5: begin
     //chislo := Copy(DateToStr(Now),1,2);
     //mes := Months[MonthOf(Now)];
     //god := Copy(DateToStr(Now),7,4);
     //S := 'Военная тайна'+' '+chislo+' '+mes+' '+god ;
     //S := 'https://www.youtube.com/results?search_query=' + S;
     //S := TNetEncoding.URL.Encode(S);
     S := 'https://foks-donetsk.com/' + S;
     Perehod(S);
   end;

   // Техносити
6: begin
     //chislo := Copy(DateToStr(Now),1,2);
     //mes := Months[MonthOf(Now)];
     //god := Copy(DateToStr(Now),7,4);
     //S := 'Легенды и мифы'+' '+chislo+' '+mes+' '+god ;
     //S := 'https://www.youtube.com/results?search_query=' + S;
     //S := TNetEncoding.URL.Encode(S);
     S := 'https://technosity-dnr.ru/' + S;
     Perehod(S);
   end;

   // Техника № 1
7: begin
     //chislo := Copy(DateToStr(Now),1,2);
     //mes := Months[MonthOf(Now)];
     //god := Copy(DateToStr(Now),7,4);
     //S := 'Неизвестная история'+' '+chislo+' '+mes+' '+god ;
     //S := 'https://www.youtube.com/results?search_query=' + S;
     //S := TNetEncoding.URL.Encode(S);
     S := 'https://xn----7sbenpmt2bhx.xn--p1ai/' + S;
     Perehod(S);
   end;

   // Бери
8: begin
     //chislo := Copy(DateToStr(Now),1,2);
     //mes := Months[MonthOf(Now)];
     //god := Copy(DateToStr(Now),7,4);
     //S := 'Совбез'+' '+chislo+' '+mes+' '+god ;
     //S := 'https://www.youtube.com/results?search_query=' + S;
     //S := TNetEncoding.URL.Encode(S);
     S := 'https://beri.shop/' + S;
     Perehod(S);
   end;
   // Ozon
9: begin
     //chislo := Copy(DateToStr(Now),1,2);
     //mes := Months[MonthOf(Now)];
     //god := Copy(DateToStr(Now),7,4);
     //S := 'Секретные материалы'+' '+chislo+' '+mes+' '+god ;
     //S := 'https://www.youtube.com/results?search_query=' + S;
     //S := TNetEncoding.URL.Encode(S);
     S := 'https://www.ozon.ru/' + S;
     Perehod(S);
   end;
   // Aliexpress
10: begin
     //chislo := Copy(DateToStr(Now),1,2);
     //mes := Months[MonthOf(Now)];
     //god := Copy(DateToStr(Now),7,4);
     //S := ''+' '+chislo+' '+mes+' '+god ;
     //S := 'https://www.youtube.com/results?search_query=' + S;
     //S := TNetEncoding.URL.Encode(S);
     S := 'https://aliexpress.ru/' + S;
     Perehod(S);
     end;
      //  Яндекс Маркет
11: begin
     //chislo := Copy(DateToStr(Now),1,2);
     //mes := Months[MonthOf(Now)];
     //god := Copy(DateToStr(Now),7,4);
     //S := ''+' '+chislo+' '+mes+' '+god ;
     //S := 'https://www.youtube.com/results?search_query=' + S;
     //S := TNetEncoding.URL.Encode(S);
     S := 'https://market.yandex.ru/' + S;
     Perehod(S);
     end;
      //Техникс
12: begin
     //chislo := Copy(DateToStr(Now),1,2);
     //mes := Months[MonthOf(Now)];
     //god := Copy(DateToStr(Now),7,4);
     //S := ''+' '+chislo+' '+mes+' '+god ;
     //S := 'https://www.youtube.com/results?search_query=' + S;
     //S := TNetEncoding.URL.Encode(S);
     S := 'https://tehniksdn.ru/' + S;
     Perehod(S);
    end;
      //Hitech
13: begin
     //chislo := Copy(DateToStr(Now),1,2);
     //mes := Months[MonthOf(Now)];
     //god := Copy(DateToStr(Now),7,4);
     //S := ''+' '+chislo+' '+mes+' '+god ;
     //S := 'https://www.youtube.com/results?search_query=' + S;
     //S := TNetEncoding.URL.Encode(S);
     S := 'https://hitech.shop/' + S;
     Perehod(S);
    end;
      //Uni-MarketPlace
14: begin
     //chislo := Copy(DateToStr(Now),1,2);
     //mes := Months[MonthOf(Now)];
     //god := Copy(DateToStr(Now),7,4);
     //S := ''+' '+chislo+' '+mes+' '+god ;
     //S := 'https://www.youtube.com/results?search_query=' + S;
     //S := TNetEncoding.URL.Encode(S);
     S := 'https://uni-marketplace.com/' + S;
     Perehod(S);
    end;
        //Технокасса
15: begin
     //chislo := Copy(DateToStr(Now),1,2);
     //mes := Months[MonthOf(Now)];
     //god := Copy(DateToStr(Now),7,4);
     //S := ''+' '+chislo+' '+mes+' '+god ;
     //S := 'https://www.youtube.com/results?search_query=' + S;
     //S := TNetEncoding.URL.Encode(S);
     S := 'https://tehkasa.ru/' + S;
     Perehod(S);
    end;




  Else
    ShellExecute(Application.Handle, 'open', '', nil, nil,SW_SHOWNOACTIVATE);
  end;

end;

procedure TForm1.ComboBox3Select(Sender: TObject);
var
it       : integer;
begin
it := Form1.ComboBox3.ItemIndex;
case it of
//
0:   ShellExecute(Application.Handle, 'open', 'https://www.gismeteo.ru/weather-donetsk-5080/?ysclid=lqh27cs1ed358502431', nil, nil,SW_SHOWNOACTIVATE);
//
1:   ShellExecute(Application.Handle, 'open', 'https://yandex.ru/pogoda/ru?ysclid=lqh28h41b1950482231&lat=47.9789&lon=37.9138', nil, nil,SW_SHOWNOACTIVATE);
//https://yandex.ru/pogoda/ru?ysclid=lqh28h41b1950482231&lat=47.9789&lon=37.9138
2:   ShellExecute(Application.Handle, 'open', 'https://pogoda7.ru/prognoz/gorod209591-Russia-DNR-Donetsk/10days/full', nil, nil,SW_SHOWNOACTIVATE);
//
3:   ShellExecute(Application.Handle, 'open', 'https://pogoda.mail.ru/prognoz/donetsk/24hours/', nil, nil,SW_SHOWNOACTIVATE);
////4:   ShellExecute(Application.Handle, 'open', 'https://rp5.ru/%D0%9F%D0%BE%D0%B3%D0%BE%D0%B4%D0%B0_%D0%B2_%D0%94%D0%BE%D0%BD%D0%B5%D1%86%D0%BA%D0%B5,_%D0%94%D0%BE%D0%BD%D0%B5%D1%86%D0%BA%D0%B0%D1%8F_%D0%9D%D0%B0%D1%80%D0%BE%D0%B4%D0%BD%D0%B0%D1%8F_%D0%A0%D0%B5%D1%81%D0%BF%D1%83%D0%B1%D0%BB%D0%B8%D0%BA%D0%B0?ysclid=lqh2bfxdl2173757901', nil, nil,SW_SHOWNOACTIVATE);

  Else
    ShellExecute(Application.Handle, 'open', 'http://www.vsetv.com/', nil, nil,SW_SHOWNOACTIVATE);
  end;

end;

procedure TForm1.ComboBox4Select(Sender: TObject);
var
it       : integer;
begin
it := Form1.ComboBox4.ItemIndex;
  case it of
   //RuTracker.org
   0:   ShellExecute(Application.Handle, 'open', 'https://rutracker.org/forum/index.php', nil, nil,SW_SHOWNOACTIVATE);
   //riperam.org
   1:   ShellExecute(Application.Handle, 'open', 'https://riper.online/', nil, nil,SW_SHOWNOACTIVATE);
   // www.bigfangroup.org
   2:   ShellExecute(Application.Handle, 'open', 'www.bigfangroup.org', nil, nil,SW_SHOWNOACTIVATE);
   // rufilmtv.tv
   3:   ShellExecute(Application.Handle, 'open', 'https://rufilmtv.tv/', nil, nil,SW_SHOWNOACTIVATE);
   // windows64.net
   4:   ShellExecute(Application.Handle, 'open', 'https://windows64.net/', nil, nil,SW_SHOWNOACTIVATE);
   //  Matrix
   5:   ShellExecute(Application.Handle, 'open', 'https://matrix.online/account', nil, nil,SW_SHOWNOACTIVATE);
   //
   6:   ShellExecute(Application.Handle, 'open', 'https://aliexpress.ru/', nil, nil,SW_SHOWNOACTIVATE);

  Else
    ShellExecute(Application.Handle, 'open', '', nil, nil,SW_SHOWNOACTIVATE);
  end;
end;

procedure TForm1.ComboBox5Select(Sender: TObject);
var
it       : integer;
begin
it := Form1.ComboBox5.ItemIndex;
case it of
//
0:   ShellExecute(Application.Handle, 'open', 'https://world-weather.ru/pogoda/russia/donetsk/sunrise/?ysclid=lb4qg3rfwe370408491', nil, nil,SW_SHOWNOACTIVATE);
//
1:   ShellExecute(Application.Handle, 'open', 'https://ru.meteocast.net/sunrise-sunset/ua/donetsk/', nil, nil,SW_SHOWNOACTIVATE);
//
2:   ShellExecute(Application.Handle, 'open', 'https://time-in.ru/sun/donetsk?ysclid=lb4rlrbdvi595900714', nil, nil,SW_SHOWNOACTIVATE);
//
3:   ShellExecute(Application.Handle, 'open', 'https://voshod-solnca.ru/equinoxes/%D0%B4%D0%BE%D0%BD%D0%B5%D1%86%D0%BA_(%D1%83%D0%BA%D1%80%D0%B0%D0%B8%D0%BD%D0%B0)', nil, nil,SW_SHOWNOACTIVATE);
//https://voshod-solnca.ru/equinoxes/%D0%B4%D0%BE%D0%BD%D0%B5%D1%86%D0%BA_(%D1%83%D0%BA%D1%80%D0%B0%D0%B8%D0%BD%D0%B0)/NaN
  Else
    ShellExecute(Application.Handle, 'open', 'http://www.vsetv.com/', nil, nil,SW_SHOWNOACTIVATE);
  end;


end;

procedure TForm1.ComboBox7Select(Sender: TObject);
var
it       : integer;
begin
it := Form1.ComboBox7.ItemIndex;
case it of
//
0:   ShellExecute(Application.Handle, 'open', 'https://my-calend.ru/name-days/today', nil, nil,SW_SHOWNOACTIVATE);
//
1:   ShellExecute(Application.Handle, 'open', 'https://nashiimena.ru/imena-v-religii/imeniny-segodnya/?ysclid=lb4uij0mk1300771630', nil, nil,SW_SHOWNOACTIVATE);
//
2:   ShellExecute(Application.Handle, 'open', 'https://www.calend.ru/names/?ysclid=lb4uhofkl826131456', nil, nil,SW_SHOWNOACTIVATE);
  Else
    ShellExecute(Application.Handle, 'open', 'http://www.vsetv.com/', nil, nil,SW_SHOWNOACTIVATE);
  end;

end;

procedure TForm1.ComboBox8Select(Sender: TObject);
var
it       : integer;
begin
it := Form1.ComboBox8.ItemIndex;
case it of
//
0:   ShellExecute(Application.Handle, 'open', 'https://kakoysegodnyaprazdnik.ru/', nil, nil,SW_SHOWNOACTIVATE);
//
1:   ShellExecute(Application.Handle, 'open', 'https://calendar.online.ua/', nil, nil,SW_SHOWNOACTIVATE);
//
2:   ShellExecute(Application.Handle, 'open', '', nil, nil,SW_SHOWNOACTIVATE);
  Else
    ShellExecute(Application.Handle, 'open', 'http://www.vsetv.com/', nil, nil,SW_SHOWNOACTIVATE);
  end;

end;

procedure TForm1.ComboBox9Select(Sender: TObject);
var
it       : integer;
begin
it := Form1.ComboBox9.ItemIndex;
case it of
//
0:   ShellExecute(Application.Handle, 'open', 'https://horo.mail.ru/prediction/gemini/today/', nil, nil,SW_SHOWNOACTIVATE);
//
1:   ShellExecute(Application.Handle, 'open', 'https://1001goroskop.ru/?znak=gemini', nil, nil,SW_SHOWNOACTIVATE);
//
2:   ShellExecute(Application.Handle, 'open', 'https://goroskop365.ru/gemini/', nil, nil,SW_SHOWNOACTIVATE);
//
3:   ShellExecute(Application.Handle, 'open', 'https://www.astrostar.ru/horoscopes/main/bliznetsi/day.html', nil, nil,SW_SHOWNOACTIVATE);
//
4:   ShellExecute(Application.Handle, 'open', 'https://gadalkindom.ru/goroskop/bliznetsy', nil, nil,SW_SHOWNOACTIVATE);
  Else
    ShellExecute(Application.Handle, 'open', 'http://www.vsetv.com/', nil, nil,SW_SHOWNOACTIVATE);
  end;

end;

procedure TForm1.ComboBoxEx1Change(Sender: TObject);
begin
//ComboBoxEx1Select(Sender);
end;

procedure TForm1.ComboBoxEx1Select(Sender: TObject);

var
it       : integer;
//ispFlag  : boolean;
begin
//if ispFlag = true then
//   begin
//     Exit;
//   end;

it := Form1.ComboBoxEx1.ItemIndex;
//ShowMessage((Sender as TForm).Name)
case it of
//Звезда
0:   ShellExecute(Application.Handle, 'open', 'http://www.vsetv.com/schedule_channel_232_week.html', nil, nil,SW_SHOWNOACTIVATE);
//Золотая коллекция
1:   ShellExecute(Application.Handle, 'open', 'http://www.vsetv.com/schedule_channel_1572_week.html', nil, nil,SW_SHOWNOACTIVATE);
//Россия 1
2:   ShellExecute(Application.Handle, 'open', 'http://www.vsetv.com/schedule_channel_9_week.html', nil, nil,SW_SHOWNOACTIVATE);
//Россия 24
3:   ShellExecute(Application.Handle, 'open', 'http://www.vsetv.com/schedule_channel_328_week.html', nil, nil,SW_SHOWNOACTIVATE);
//5 канал
4:   ShellExecute(Application.Handle, 'open', 'http://www.vsetv.com/schedule_channel_308_week.html', nil, nil,SW_SHOWNOACTIVATE);
//Amedia
5:   ShellExecute(Application.Handle, 'open', 'http://www.vsetv.com/schedule_channel_895_week.html', nil, nil,SW_SHOWNOACTIVATE);
//Дом кино Премиум HD
6:   ShellExecute(Application.Handle, 'open', 'http://www.vsetv.com/schedule_channel_1429_week.html', nil, nil,SW_SHOWNOACTIVATE);
//Дом кино
7:   ShellExecute(Application.Handle, 'open', 'http://www.vsetv.com/schedule_channel_302_week.html', nil, nil,SW_SHOWNOACTIVATE);
//Известия
8:   ShellExecute(Application.Handle, 'open', 'https://federal.tv/program/izru?ysclid=l8h6jsyxgn949004620', nil, nil,SW_SHOWNOACTIVATE);
//Иллюзион+
9:   ShellExecute(Application.Handle, 'open', 'http://www.vsetv.com/schedule_channel_327_week.html', nil, nil,SW_SHOWNOACTIVATE);
//Русский иллюзион
10:   ShellExecute(Application.Handle, 'open', 'http://www.vsetv.com/schedule_channel_177_week.html', nil, nil,SW_SHOWNOACTIVATE);
//Кинокомедия
11:   ShellExecute(Application.Handle, 'open', 'http://www.vsetv.com/schedule_channel_292_week.html', nil, nil,SW_SHOWNOACTIVATE);
//Русская комедия
12:   ShellExecute(Application.Handle, 'open', 'http://www.vsetv.com/schedule_channel_1081_week.html', nil, nil,SW_SHOWNOACTIVATE);
//Киномикс
13:   ShellExecute(Application.Handle, 'open', 'http://www.vsetv.com/schedule_channel_608_week.html', nil, nil,SW_SHOWNOACTIVATE);
//Киносвидание
14:   ShellExecute(Application.Handle, 'open', 'http://www.vsetv.com/schedule_channel_1665_week.html', nil, nil,SW_SHOWNOACTIVATE);
//Киносемья
15:   ShellExecute(Application.Handle, 'open', 'http://www.vsetv.com/schedule_channel_1692_week.html', nil, nil,SW_SHOWNOACTIVATE);
//Киносерия
16:   ShellExecute(Application.Handle, 'open', 'http://www.vsetv.com/schedule_channel_278_week.html', nil, nil,SW_SHOWNOACTIVATE);
//Кинохит
17:   ShellExecute(Application.Handle, 'open', 'http://www.vsetv.com/schedule_channel_1693_week.html', nil, nil,SW_SHOWNOACTIVATE);
//Кто есть кто
18:   ShellExecute(Application.Handle, 'open', 'http://www.vsetv.com/schedule_channel_399_week.html', nil, nil,SW_SHOWNOACTIVATE);
//Любимое кино
19:   ShellExecute(Application.Handle, 'open', 'http://www.vsetv.com/schedule_channel_1195_week.html', nil, nil,SW_SHOWNOACTIVATE);
//Мир сериала
20:   ShellExecute(Application.Handle, 'open', 'http://www.vsetv.com/schedule_channel_637_week.html', nil, nil,SW_SHOWNOACTIVATE);
//Мужское кино
21:   ShellExecute(Application.Handle, 'open', 'http://www.vsetv.com/schedule_channel_1060_week.html', nil, nil,SW_SHOWNOACTIVATE);
//Мужской
22:   ShellExecute(Application.Handle, 'open', 'http://www.vsetv.com/schedule_channel_648_week.html', nil, nil,SW_SHOWNOACTIVATE);
//НТВ
23:   ShellExecute(Application.Handle, 'open', 'http://www.vsetv.com/schedule_channel_10_week.html', nil, nil,SW_SHOWNOACTIVATE);
//НТВ Право
24:   ShellExecute(Application.Handle, 'open', 'http://www.vsetv.com/schedule_channel_1215_week.html', nil, nil,SW_SHOWNOACTIVATE);
//НТВ Сериал
25:   ShellExecute(Application.Handle, 'open', 'http://www.vsetv.com/schedule_channel_1214_week.html', nil, nil,SW_SHOWNOACTIVATE);
//Наше новое кино
26:   ShellExecute(Application.Handle, 'open', 'http://www.vsetv.com/schedule_channel_260_week.html', nil, nil,SW_SHOWNOACTIVATE);
//Новороссия
27:   ShellExecute(Application.Handle, 'open', '', nil, nil,SW_SHOWNOACTIVATE);
//Ностальгия
28:   ShellExecute(Application.Handle, 'open', 'http://www.vsetv.com/schedule_channel_216_week.html', nil, nil,SW_SHOWNOACTIVATE);
//Оплот 2
29:   ShellExecute(Application.Handle, 'open', 'https://vsednr.ru/teleprogramma-oplot-2/', nil, nil,SW_SHOWNOACTIVATE);
//Оружие
30:   ShellExecute(Application.Handle, 'open', 'http://www.vsetv.com/schedule_channel_708_week.html', nil, nil,SW_SHOWNOACTIVATE);
//Первый
31:   ShellExecute(Application.Handle, 'open', 'http://www.vsetv.com/schedule_channel_8_week.html', nil, nil,SW_SHOWNOACTIVATE);
//Первый космический
32:   ShellExecute(Application.Handle, 'open', 'http://www.vsetv.com/schedule_channel_970_week.html', nil, nil,SW_SHOWNOACTIVATE);
//Победа
33:   ShellExecute(Application.Handle, 'open', 'http://www.vsetv.com/schedule_channel_1425_week.html', nil, nil,SW_SHOWNOACTIVATE);
//Пятница
34:   ShellExecute(Application.Handle, 'open', 'http://www.vsetv.com/schedule_channel_48_week.html', nil, nil,SW_SHOWNOACTIVATE);
//Родное кино
35:   ShellExecute(Application.Handle, 'open', 'http://www.vsetv.com/schedule_channel_312_week.html', nil, nil,SW_SHOWNOACTIVATE);
//Русская комедия
36:   ShellExecute(Application.Handle, 'open', '', nil, nil,SW_SHOWNOACTIVATE);
//Русский бестселлер
37:   ShellExecute(Application.Handle, 'open', 'http://www.vsetv.com/schedule_channel_812_week.html', nil, nil,SW_SHOWNOACTIVATE);
//Русский детектив
38:   ShellExecute(Application.Handle, 'open', 'http://www.vsetv.com/schedule_channel_942_week.html', nil, nil,SW_SHOWNOACTIVATE);
//Русский роман
39:   ShellExecute(Application.Handle, 'open', 'http://www.vsetv.com/schedule_channel_757_week.html', nil, nil,SW_SHOWNOACTIVATE);
//Спас
40:   ShellExecute(Application.Handle, 'open', 'http://www.vsetv.com/schedule_channel_238_week.html', nil, nil,SW_SHOWNOACTIVATE);
//Суббота
41:   ShellExecute(Application.Handle, 'open', 'http://www.vsetv.com/schedule_channel_1302_week.html', nil, nil,SW_SHOWNOACTIVATE);
//Т 24
42:   ShellExecute(Application.Handle, 'open', 'http://www.vsetv.com/schedule_channel_381_week.html', nil, nil,SW_SHOWNOACTIVATE);
//ТВ 3
43:   ShellExecute(Application.Handle, 'open', 'http://www.vsetv.com/schedule_channel_63_week.html', nil, nil,SW_SHOWNOACTIVATE);
//Юнион
44:   ShellExecute(Application.Handle, 'open', '', nil, nil,SW_SHOWNOACTIVATE);
//Discovery Channel
45:   ShellExecute(Application.Handle, 'open', 'http://www.vsetv.com/schedule_channel_38_week.html', nil, nil,SW_SHOWNOACTIVATE);
//Discovery Science
46:   ShellExecute(Application.Handle, 'open', 'http://www.vsetv.com/schedule_channel_41_week.html', nil, nil,SW_SHOWNOACTIVATE);
//National Geographic
47:   ShellExecute(Application.Handle, 'open', 'http://www.vsetv.com/schedule_channel_594_week.html', nil, nil,SW_SHOWNOACTIVATE);
//Ren TV
48:   ShellExecute(Application.Handle, 'open', 'http://www.vsetv.com/schedule_channel_13_week.html', nil, nil,SW_SHOWNOACTIVATE);
//Viasat History
49:   ShellExecute(Application.Handle, 'open', 'http://www.vsetv.com/schedule_channel_1505_week.html', nil, nil,SW_SHOWNOACTIVATE);
//Viasat Nature/History HD
50:   ShellExecute(Application.Handle, 'open', 'http://www.vsetv.com/schedule_channel_1505_week.html', nil, nil,SW_SHOWNOACTIVATE);
//TV1000
51:   ShellExecute(Application.Handle, 'open', 'http://www.vsetv.com/schedule_channel_107_week.html', nil, nil,SW_SHOWNOACTIVATE);
//TV1000 Action
52:   ShellExecute(Application.Handle, 'open', 'http://www.vsetv.com/schedule_channel_458_week.html', nil, nil,SW_SHOWNOACTIVATE);
//TV1000 Русское кино
53:   ShellExecute(Application.Handle, 'open', 'http://www.vsetv.com/schedule_channel_255_week.html', nil, nil,SW_SHOWNOACTIVATE);
//Оплот 2
54:   ShellExecute(Application.Handle, 'open', 'https://vsednr.ru/teleprogramma-oplot-2/', nil, nil,SW_SHOWNOACTIVATE);
//Оплот ТВ
55:   ShellExecute(Application.Handle, 'open', 'https://telekanal-oplot.tv/teleprogramma/?ysclid=l8h2gxtwv8105996348', nil, nil,SW_SHOWNOACTIVATE);

  Else
    ShellExecute(Application.Handle, 'open', 'http://www.vsetv.com/', nil, nil,SW_SHOWNOACTIVATE);
  end;
  //ispFlag := not ispFlag;
end;


procedure TForm1.ComboBoxEx2KeyPress(Sender: TObject; var Key: Char);
begin
if key = #13 then
   begin
     Button16Click(Sender);
   end;

end;

procedure TForm1.ComboBoxEx4Select(Sender: TObject);
begin
Form1.Backup_drive := Form1.ComboBoxEx4.Text;
end;


procedure TForm1.DateTimePicker5Change(Sender: TObject);
begin
//
Form1.SpinEdit5.Enabled := false;
end;

procedure TForm1.DBGrid2DblClick(Sender: TObject);
var
id : integer;
begin
//
//sel_col :=Form1.DBGrid2.SelectedIndex; //определяем номер выделенного поля (столбца)
//sel_row := DataModule1.DataSource_osn.DataSet.RecNo; // номер строки

id:=Datamodule1.DataSource_osn.DataSet.FieldByName('id').AsInteger;
sel_row := id;
if sel_row <> 0 then
  begin
   Form9.Memo1.Lines.Clear;
   Form9.kkk := Form1.sel_row ;
   Form9.Text_sql := 'select * from Bill_New where id = :p1';
   Form9.FDQuery1.Active := false;
   Form9.FDQuery1.SQL.Clear;
   Form9.FDQuery1.SQL.Add(Form9.Text_SQL);
   Form9.FDQuery1.Params.ParamByName('p1').Value :=Form9.kkk ;
   Form9.FDQuery1.Active :=true;

   Form9.zamet := Form9.FDQuery1.FieldByName('Zam').AsString;
   Form9.Memo1.Text := Form9.zamet;
   Form9.s := 'Изменение заметок за  '
   + Form9.FDQuery1.FieldByName('Data').AsString
   + '  число';
Form9.Caption := Form9.s;
//Form9.Memo1.SetFocus  ;

   Form9.Show;
   //Form1.Hide;
  end;
end;

procedure TForm1.DBGrid2DrawColumnCell(Sender: TObject; const Rect: TRect;
  DataCol: Integer; Column: TColumn; State: TGridDrawState);
begin

//=====================================================================
 if Column.FieldName = 'DenNed' then //Удалите эту линию, если хотете закрасить целую линию//
    if Column.Field.Dataset.FieldByName('DenNed').AsString <> '' then
      if  (gdFocused in State) then //имеет ли ячейка фокус?
         DBGrid2.Canvas.Brush.Color := clBlack //имеет фокус
      else
        DBGrid2.Canvas.Brush.Color :=   RGB(244, 176, 132); //не имеет фокуса
 //Теперь давайте закрасим ячейку используя стандартный метод:
  DBGrid2.DefaultDrawColumnCell(Rect, DataCol, Column, State);
//=====================================================================
 if Column.FieldName = 'DenNed' then //Удалите эту линию, если хотете закрасить целую линию//
    if Column.Field.Dataset.FieldByName('DenNed').AsString = 'Пятница' then
      if  (gdFocused in State) then //имеет ли ячейка фокус?
         DBGrid2.Canvas.Brush.Color := clBlack //имеет фокус
      else
        DBGrid2.Canvas.Brush.Color :=   RGB(244, 176, 232); //не имеет фокуса
 //Теперь давайте закрасим ячейку используя стандартный метод:
  DBGrid2.DefaultDrawColumnCell(Rect, DataCol, Column, State);
//=====================================================================
 if Column.FieldName = 'DenNed' then //Удалите эту линию, если хотете закрасить целую линию//
    if Column.Field.Dataset.FieldByName('DenNed').AsString = 'Суббота' then
      if  (gdFocused in State) then //имеет ли ячейка фокус?
         DBGrid2.Canvas.Brush.Color := clBlack //имеет фокус
      else
        DBGrid2.Canvas.Brush.Color :=   RGB(54, 202, 79); //не имеет фокуса
 //Теперь давайте закрасим ячейку используя стандартный метод:
  DBGrid2.DefaultDrawColumnCell(Rect, DataCol, Column, State);
//=====================================================================
 if Column.FieldName = 'DenNed' then //Удалите эту линию, если хотете закрасить целую линию//
    if Copy(Column.Field.Dataset.FieldByName('DenNed').AsString,1,5) = 'Воскр' then
      if  (gdFocused in State) then //имеет ли ячейка фокус?
         DBGrid2.Canvas.Brush.Color := clBlack //имеет фокус
      else
        DBGrid2.Canvas.Brush.Color :=   RGB(90, 83, 224); //не имеет фокуса
 //Теперь давайте закрасим ячейку используя стандартный метод:
  DBGrid2.DefaultDrawColumnCell(Rect, DataCol, Column, State);
//=====================================================================
 if Column.FieldName = 'Travm' then //Удалите эту линию, если хотете закрасить целую линию//
    if Column.Field.Dataset.FieldByName('Travm').AsString <> '' then
      if (gdFocused in State) then //имеет ли ячейка фокус?
         DBGrid2.Canvas.Brush.Color := clBlack //имеет фокус
      else
        DBGrid2.Canvas.Brush.Color := RGB(189, 215, 238); //не имеет фокуса
 //Теперь давайте закрасим ячейку используя стандартный метод:
  DBGrid2.DefaultDrawColumnCell(Rect, DataCol, Column, State);
//=====================================================================
 if Column.FieldName = 'Gips' then //Удалите эту линию, если хотете закрасить целую линию//
    if Column.Field.Dataset.FieldByName('Gips').AsString <> '' then
      if (gdFocused in State) then //имеет ли ячейка фокус?
         DBGrid2.Canvas.Brush.Color := clBlack //имеет фокус
      else
        DBGrid2.Canvas.Brush.Color := RGB(155, 194, 230); //не имеет фокуса
 //Теперь давайте закрасим ячейку используя стандартный метод:
  DBGrid2.DefaultDrawColumnCell(Rect, DataCol, Column, State);
//=====================================================================
 if Column.FieldName = 'Sobitie' then //Удалите эту линию, если хотете закрасить целую линию//
    if (Column.Field.Dataset.FieldByName('Sobitie').AsString = 'Псрл')
    or (Column.Field.Dataset.FieldByName('Sobitie').AsString = 'псрл')
    or (Column.Field.Dataset.FieldByName('Sobitie').AsString = 'прcл') then

      if (gdFocused in State) then //имеет ли ячейка фокус?
         DBGrid2.Canvas.Brush.Color := clBlack //имеет фокус
      else
        DBGrid2.Canvas.Brush.Color := RGB(0, 176, 240); //не имеет фокуса
 //Теперь давайте закрасим ячейку используя стандартный метод:
  DBGrid2.DefaultDrawColumnCell(Rect, DataCol, Column, State);
//=====================================================================
//=====================================================================
 if Column.FieldName = 'Sobitie' then //Удалите эту линию, если хотете закрасить целую линию//
    if (Copy(Column.Field.Dataset.FieldByName('Sobitie').AsString,1,3) = 'З/п')
    then

      if (gdFocused in State) then //имеет ли ячейка фокус?
         DBGrid2.Canvas.Brush.Color := clBlack //имеет фокус
      else
        DBGrid2.Canvas.Brush.Color := RGB(196, 60, 70); //не имеет фокуса
 //Теперь давайте закрасим ячейку используя стандартный метод:
  DBGrid2.DefaultDrawColumnCell(Rect, DataCol, Column, State);
//=====================================================================
//=====================================================================
 if Column.FieldName = 'Pil' then //Удалите эту линию, если хотете закрасить целую линию//
    if Column.Field.Dataset.FieldByName('Pil').AsString = 'НЕ ПИЛ' then
      if (gdFocused in State) then //имеет ли ячейка фокус?
         DBGrid2.Canvas.Brush.Color := clBlack //имеет фокус
      else
        DBGrid2.Canvas.Brush.Color := RGB(50, 193, 47); //не имеет фокуса
 //Теперь давайте закрасим ячейку используя стандартный метод:
  DBGrid2.DefaultDrawColumnCell(Rect, DataCol, Column, State);
//=====================================================================
//=====================================================================
 if Column.FieldName = 'Pil' then //Удалите эту линию, если хотете закрасить целую линию//
    if Column.Field.Dataset.FieldByName('Pil').AsString = 'ПИЛ' then
      if (gdFocused in State) then //имеет ли ячейка фокус?
         DBGrid2.Canvas.Brush.Color := clBlack //имеет фокус
      else
        DBGrid2.Canvas.Brush.Color := RGB(255, 0, 0); //не имеет фокуса
 //Теперь давайте закрасим ячейку используя стандартный метод:
  DBGrid2.DefaultDrawColumnCell(Rect, DataCol, Column, State);
//=====================================================================
//=====================================================================
 if Column.FieldName = 'Zam' then //Удалите эту линию, если хотете закрасить целую линию//
    if Column.Field.Dataset.FieldByName('Zam').AsString <> '' then
      if (gdFocused in State) then //имеет ли ячейка фокус?
         DBGrid2.Canvas.Brush.Color := clBlack //имеет фокус
      else
        DBGrid2.Canvas.Brush.Color := RGB(255, 0, 0); //не имеет фокуса
 //Теперь давайте закрасим ячейку используя стандартный метод:
  DBGrid2.DefaultDrawColumnCell(Rect, DataCol, Column, State);
//=====================================================================

end;

procedure TForm1.DBGrid3DrawColumnCell(Sender: TObject; const Rect: TRect;
  DataCol: Integer; Column: TColumn; State: TGridDrawState);
begin

//=====================================================================
 if Column.FieldName = 'DenNed' then //Удалите эту линию, если хотете закрасить целую линию//
    if Column.Field.Dataset.FieldByName('DenNed').AsString <> '' then
      if  (gdFocused in State) then //имеет ли ячейка фокус?
         DBGrid3.Canvas.Brush.Color := clBlack //имеет фокус
      else
        DBGrid3.Canvas.Brush.Color :=   RGB(244, 176, 132); //не имеет фокуса
 //Теперь давайте закрасим ячейку используя стандартный метод:
  DBGrid3.DefaultDrawColumnCell(Rect, DataCol, Column, State);
//=====================================================================
 if Column.FieldName = 'DenNed' then //Удалите эту линию, если хотете закрасить целую линию//
    if Column.Field.Dataset.FieldByName('DenNed').AsString = 'Пятница' then
      if  (gdFocused in State) then //имеет ли ячейка фокус?
         DBGrid3.Canvas.Brush.Color := clBlack //имеет фокус
      else
        DBGrid3.Canvas.Brush.Color :=   RGB(244, 176, 232); //не имеет фокуса
 //Теперь давайте закрасим ячейку используя стандартный метод:
  DBGrid3.DefaultDrawColumnCell(Rect, DataCol, Column, State);
//=====================================================================
 if Column.FieldName = 'DenNed' then //Удалите эту линию, если хотете закрасить целую линию//
    if Column.Field.Dataset.FieldByName('DenNed').AsString = 'Суббота' then
      if  (gdFocused in State) then //имеет ли ячейка фокус?
         DBGrid3.Canvas.Brush.Color := clBlack //имеет фокус
      else
        DBGrid3.Canvas.Brush.Color :=   RGB(54, 202, 79); //не имеет фокуса
 //Теперь давайте закрасим ячейку используя стандартный метод:
  DBGrid3.DefaultDrawColumnCell(Rect, DataCol, Column, State);
//=====================================================================
 if Column.FieldName = 'DenNed' then //Удалите эту линию, если хотете закрасить целую линию//
    if Copy(Column.Field.Dataset.FieldByName('DenNed').AsString,1,5) = 'Воскр' then
      if  (gdFocused in State) then //имеет ли ячейка фокус?
         DBGrid3.Canvas.Brush.Color := clBlack //имеет фокус
      else
        DBGrid3.Canvas.Brush.Color :=   RGB(90, 83, 224); //не имеет фокуса
 //Теперь давайте закрасим ячейку используя стандартный метод:
  DBGrid3.DefaultDrawColumnCell(Rect, DataCol, Column, State);
//=====================================================================

//=====================================================================
 if Column.FieldName = 'Travm' then //Удалите эту линию, если хотете закрасить целую линию//
    if Column.Field.Dataset.FieldByName('Travm').AsString <> '' then
      if (gdFocused in State) then //имеет ли ячейка фокус?
         DBGrid3.Canvas.Brush.Color := clBlack //имеет фокус
      else
        DBGrid3.Canvas.Brush.Color := RGB(189, 215, 238); //не имеет фокуса
 //Теперь давайте закрасим ячейку используя стандартный метод:
  DBGrid3.DefaultDrawColumnCell(Rect, DataCol, Column, State);
//=====================================================================
 if Column.FieldName = 'Gips' then //Удалите эту линию, если хотете закрасить целую линию//
    if Column.Field.Dataset.FieldByName('Gips').AsString <> '' then
      if (gdFocused in State) then //имеет ли ячейка фокус?
         DBGrid3.Canvas.Brush.Color := clBlack //имеет фокус
      else
        DBGrid3.Canvas.Brush.Color := RGB(155, 194, 230); //не имеет фокуса
 //Теперь давайте закрасим ячейку используя стандартный метод:
  DBGrid3.DefaultDrawColumnCell(Rect, DataCol, Column, State);
//=====================================================================
 if Column.FieldName = 'Sobitie' then //Удалите эту линию, если хотете закрасить целую линию//
    if (Column.Field.Dataset.FieldByName('Sobitie').AsString = 'Псрл')
    or (Column.Field.Dataset.FieldByName('Sobitie').AsString = 'псрл')
    or (Column.Field.Dataset.FieldByName('Sobitie').AsString = 'прcл') then

      if (gdFocused in State) then //имеет ли ячейка фокус?
         DBGrid3.Canvas.Brush.Color := clBlack //имеет фокус
      else
        DBGrid3.Canvas.Brush.Color := RGB(0, 176, 240); //не имеет фокуса
 //Теперь давайте закрасим ячейку используя стандартный метод:
  DBGrid3.DefaultDrawColumnCell(Rect, DataCol, Column, State);
//=====================================================================
//=====================================================================
 if Column.FieldName = 'Sobitie' then //Удалите эту линию, если хотете закрасить целую линию//
    if (Copy(Column.Field.Dataset.FieldByName('Sobitie').AsString,1,3) = 'З/п')
    then

      if (gdFocused in State) then //имеет ли ячейка фокус?
         DBGrid3.Canvas.Brush.Color := clBlack //имеет фокус
      else
        DBGrid3.Canvas.Brush.Color := RGB(196, 60, 70); //не имеет фокуса
 //Теперь давайте закрасим ячейку используя стандартный метод:
  DBGrid3.DefaultDrawColumnCell(Rect, DataCol, Column, State);
//=====================================================================
//=====================================================================
 if Column.FieldName = 'Pil' then //Удалите эту линию, если хотете закрасить целую линию//
    if Column.Field.Dataset.FieldByName('Pil').AsString = 'НЕ ПИЛ' then
      if (gdFocused in State) then //имеет ли ячейка фокус?
         DBGrid3.Canvas.Brush.Color := clBlack //имеет фокус
      else
        DBGrid3.Canvas.Brush.Color := RGB(50, 193, 47); //не имеет фокуса
 //Теперь давайте закрасим ячейку используя стандартный метод:
  DBGrid3.DefaultDrawColumnCell(Rect, DataCol, Column, State);
//=====================================================================
//=====================================================================
 if Column.FieldName = 'Pil' then //Удалите эту линию, если хотете закрасить целую линию//
    if Column.Field.Dataset.FieldByName('Pil').AsString = 'ПИЛ' then
      if (gdFocused in State) then //имеет ли ячейка фокус?
         DBGrid3.Canvas.Brush.Color := clBlack //имеет фокус
      else
        DBGrid3.Canvas.Brush.Color := RGB(255, 0, 0); //не имеет фокуса
 //Теперь давайте закрасим ячейку используя стандартный метод:
  DBGrid3.DefaultDrawColumnCell(Rect, DataCol, Column, State);
//=====================================================================
//=====================================================================
 if Column.FieldName = 'Zam' then //Удалите эту линию, если хотете закрасить целую линию//
    if Column.Field.Dataset.FieldByName('Zam').AsString <> '' then
      if (gdFocused in State) then //имеет ли ячейка фокус?
         DBGrid3.Canvas.Brush.Color := clBlack //имеет фокус
      else
        DBGrid3.Canvas.Brush.Color := RGB(255, 0, 0); //не имеет фокуса
 //Теперь давайте закрасим ячейку используя стандартный метод:
  DBGrid3.DefaultDrawColumnCell(Rect, DataCol, Column, State);
//=====================================================================

end;

procedure TForm1.DBGrid4DblClick(Sender: TObject);
begin
sel_row := DataModule1.DataSource_Svet.DataSet.RecNo;

Form8.Show;
//Form1.Hide;



end;

procedure TForm1.DBGrid4DrawColumnCell(Sender: TObject; const Rect: TRect;
  DataCol: Integer; Column: TColumn; State: TGridDrawState);
begin
//=====================================================================
 if Column.FieldName = 'Date_Opl' then //Удалите эту линию, если хотете закрасить целую линию//
    if Column.Field.Dataset.FieldByName('Date_Opl').AsString <> '' then
      if (gdFocused in State) then //имеет ли ячейка фокус?
         DBGrid4.Canvas.Brush.Color := clBlack //имеет фокус
      else
        DBGrid4.Canvas.Brush.Color := RGB(50, 193, 47); //не имеет фокуса
 //Теперь давайте закрасим ячейку используя стандартный метод:
  DBGrid4.DefaultDrawColumnCell(Rect, DataCol, Column, State);
//=====================================================================
 if Column.FieldName = 'Date_Opl' then //Удалите эту линию, если хотете закрасить целую линию//
    if Column.Field.Dataset.FieldByName('Date_Opl').AsString = '' then
      if (gdFocused in State) then //имеет ли ячейка фокус?
         DBGrid4.Canvas.Brush.Color := clBlack //имеет фокус
      else
        DBGrid4.Canvas.Brush.Color := RGB(255, 0, 0); //не имеет фокуса
 //Теперь давайте закрасим ячейку используя стандартный метод:
  DBGrid4.DefaultDrawColumnCell(Rect, DataCol, Column, State);
end;

procedure TForm1.DBGrid5CellClick(Column: TColumn);
var
   i:integer;
   s:string;
begin
i := Form1.DBGrid5.SelectedIndex;
s := Form1.DBGrid5.DataSource.DataSet.Fields[i].AsString ;
//sel_row  := DataModule1.DataSource_Kvartira.DataSet.RecNo;
Clipboard.AsText := s;
end;

procedure TForm1.DBGrid5DblClick(Sender: TObject);
var

Date_Oplati  : string ;
//Summa_Oplati : integer;
Dolg         : integer;
//Za_chto      : integer;
//Status       : string;
begin
//sel_row  := DataModule1.DataSource_Kvartira.DataSet.RecNo;
Date_Oplati := DataModule1.FDQuery_Kvartira.FieldByName('Date_Oplati').AsString;
Form7.SpinEdit1.Value := DataModule1.FDQuery_Kvartira.FieldByName('Summa_Oplati').AsInteger;
Form7.Edit1.Text := DataModule1.FDQuery_Kvartira.FieldByName('Za_chto').AsString;
Form7.DateTimePicker1.DateTime := DataModule1.FDQuery_Kvartira.FieldByName('Date_Oplati').AsDateTime;

Dolg := DataModule1.FDQuery_Kvartira.FieldByName('Dolg').AsInteger;
//Dolg := Dolg - Form7.SpinEdit1.Value;
Form7.Edit2.Text := DataModule1.FDQuery_Kvartira.FieldByName('Dolg').AsString;
Form7.DateTimePicker1.Date := StrToDate(today_str);
if Dolg > 0 then
   begin
    Form7.Edit2.Color := clRed;
   end;
//Date_Oplati :=
Form7.Edit3.Text := DataModule1.FDQuery_Kvartira.FieldByName('Status').AsString;
if Form7.Edit3.Text ='Оплачен' then
  begin
   Form7.Edit3.Color := clLime;
  end
else
  begin
   Form7.Edit3.Color  := clRed;
  End;

Form7.Show;
//Form1.Hide;

end;

procedure TForm1.DBGrid5DrawColumnCell(Sender: TObject; const Rect: TRect;
  DataCol: Integer; Column: TColumn; State: TGridDrawState);
begin
//=====================================================================
 if Column.FieldName = 'Status' then //Удалите эту линию, если хотете закрасить целую линию//
    if Column.Field.Dataset.FieldByName('Status').AsString = 'Оплачен' then
      if (gdFocused in State) then //имеет ли ячейка фокус?
         DBGrid5.Canvas.Brush.Color := clBlack //имеет фокус
      else
        DBGrid5.Canvas.Brush.Color := RGB(50, 193, 47); //не имеет фокуса
 //Теперь давайте закрасим ячейку используя стандартный метод:
  DBGrid5.DefaultDrawColumnCell(Rect, DataCol, Column, State);
//=====================================================================
 if Column.FieldName = 'Dolg' then //Удалите эту линию, если хотете закрасить целую линию//
    if Column.Field.Dataset.FieldByName('Dolg').AsString <> '' then
      if (gdFocused in State) then //имеет ли ячейка фокус?
         DBGrid5.Canvas.Brush.Color := clBlack //имеет фокус
      else
        DBGrid5.Canvas.Brush.Color := RGB(255, 0, 0); //не имеет фокуса
 //Теперь давайте закрасим ячейку используя стандартный метод:
  DBGrid5.DefaultDrawColumnCell(Rect, DataCol, Column, State);
//=====================================================================
 if Column.FieldName = 'Summa_Oplati' then //Удалите эту линию, если хотете закрасить целую линию//
    if Column.Field.Dataset.FieldByName('Summa_Oplati').AsString <> '' then
      if (gdFocused in State) then //имеет ли ячейка фокус?
         DBGrid5.Canvas.Brush.Color := clBlack //имеет фокус
      else
        DBGrid5.Canvas.Brush.Color := RGB(0, 102, 255); //не имеет фокуса
 //Теперь давайте закрасим ячейку используя стандартный метод:
  DBGrid5.DefaultDrawColumnCell(Rect, DataCol, Column, State);


end;

procedure TForm1.DBGrid6DrawColumnCell(Sender: TObject; const Rect: TRect;
  DataCol: Integer; Column: TColumn; State: TGridDrawState);
begin
begin
//=====================================================================
 if Column.FieldName = 'id' then //Удалите эту линию, если хотете закрасить целую линию//
    if Column.Field.Dataset.FieldByName('id').AsString <> '' then
      if  (gdFocused in State) then //имеет ли ячейка фокус?
         DBGrid6.Canvas.Brush.Color := clGradientInactiveCaption//clBlack ;//имеет фокус
      else
        DBGrid6.Canvas.Brush.Color := RGB(50, 193, 47); //не имеет фокуса
 //Теперь давайте закрасим ячейку используя стандартный метод:
  DBGrid6.DefaultDrawColumnCell(Rect, DataCol, Column, State);
         //DBGrid6.Font.Color := clWhite;  //clGradientInactiveCaption

//=====================================================================
 if Column.FieldName = 'From_Date' then //Удалите эту линию, если хотете закрасить целую линию//
    if Column.Field.Dataset.FieldByName('From_Date').AsString <> '' then
      if  (gdFocused in State) then //имеет ли ячейка фокус?
         DBGrid6.Canvas.Brush.Color := clGradientInactiveCaption //имеет фокус
      else
        DBGrid6.Canvas.Brush.Color :=   RGB(244, 176, 132); //не имеет фокуса
 //Теперь давайте закрасим ячейку используя стандартный метод:
  DBGrid6.DefaultDrawColumnCell(Rect, DataCol, Column, State);
//=====================================================================
 if Column.FieldName = 'To_Date' then //Удалите эту линию, если хотете закрасить целую линию//
    if Column.Field.Dataset.FieldByName('To_Date').AsString <> '' then
      if  (gdFocused in State) then //имеет ли ячейка фокус?
         DBGrid6.Canvas.Brush.Color := clBlack //имеет фокус
      else
        DBGrid6.Canvas.Brush.Color :=   RGB(244, 176, 132); //не имеет фокуса
 //Теперь давайте закрасим ячейку используя стандартный метод:
  DBGrid6.DefaultDrawColumnCell(Rect, DataCol, Column, State);
 //=====================================================================
 if Column.FieldName = 'Company' then //Удалите эту линию, если хотете закрасить целую линию//
    if Column.Field.Dataset.FieldByName('Company').AsString <> '' then
      if (gdFocused in State) then //имеет ли ячейка фокус?
         DBGrid6.Canvas.Brush.Color := clBlack //имеет фокус
      else
        DBGrid6.Canvas.Brush.Color := RGB(189, 215, 238); //не имеет фокуса
 //Теперь давайте закрасим ячейку используя стандартный метод:
  DBGrid6.DefaultDrawColumnCell(Rect, DataCol, Column, State);
{
//=====================================================================
 if Assigned(Column) then
  begin
   DBGrid6.Canvas.FillRect(Rect);
   DBGrid6.Canvas.TextRect(Rect, Rect.Left, Rect.Top, ' '+Column.Field.AsString);
  end;
}
 end;
end;

procedure TForm1.DBGrid7DblClick(Sender: TObject);
begin
//
sel_row := DataModule1.DataSource_Log.DataSet.RecNo; // номер строки
Form10.Show;
Form1.Hide;
Form10.Edit1.Text := DataModule1.DataSource_Log.DataSet.FieldByName('url').AsString;
Form10.Edit2.Text := DataModule1.DataSource_Log.DataSet.FieldByName('username').AsString;
Form10.Edit3.Text := DataModule1.DataSource_Log.DataSet.FieldByName('password').AsString;
//password
end;

procedure TForm1.DBGrid8DrawColumnCell(Sender: TObject; const Rect: TRect;
  DataCol: Integer; Column: TColumn; State: TGridDrawState);
begin
//=====================================================================
 if Column.FieldName = 'P_Hol' then //Удалите эту линию, если хотете закрасить целую линию//
    if Column.Field.Dataset.FieldByName('P_Hol').AsString <> '' then
      if (gdFocused in State) then //имеет ли ячейка фокус?
         DBGrid8.Canvas.Brush.Color := clBlack //имеет фокус
      else
        DBGrid8.Canvas.Brush.Color := RGB(102, 204, 255); //не имеет фокуса
 //Теперь давайте закрасим ячейку используя стандартный метод:
  DBGrid8.DefaultDrawColumnCell(Rect, DataCol, Column, State);

  //=====================================================================
 if Column.FieldName = 'P_Gor' then //Удалите эту линию, если хотете закрасить целую линию//
    if Column.Field.Dataset.FieldByName('P_Gor').AsString <> '' then
      if (gdFocused in State) then //имеет ли ячейка фокус?
         DBGrid8.Canvas.Brush.Color := clBlack //имеет фокус
      else
        DBGrid8.Canvas.Brush.Color := RGB(255, 51, 204); //не имеет фокуса
 //Теперь давайте закрасим ячейку используя стандартный метод:
  DBGrid8.DefaultDrawColumnCell(Rect, DataCol, Column, State);

end;

procedure TForm1.DBGrid9CellClick(Column: TColumn);
var
   i:integer;
   s:string;
begin
i :=Form1.DBGrid9.SelectedIndex;
 s := Form1.DBGrid9.DataSource.DataSet.Fields[i].AsString ;
Form1.Edit17.Text := s;


DataModule1.FDQuery_EPG_CUR.Active := false;
DataModule1.FDQuery_EPG_CUR.Active := true;
end;

procedure TForm1.DBGrid9DrawColumnCell(Sender: TObject; const Rect: TRect;
  DataCol: Integer; Column: TColumn; State: TGridDrawState);
var
 rect2_width :integer;
 p : integer;//-->X
 s : string;
 //my var
myHour, myMin, mySec, myMilli : word;
t_start : TdateTime;
t_stop  : TdateTime;
s_nach  : TdateTime;
seychas : TdateTime;
prodolg :integer;
a,b,c :integer;
x     :Double;
begin
//
if Column.FieldName = 'Progress' then
   begin
   DecodeTime(Time,myHour, myMin, mySec, myMilli);
   t_start :=DataModule1.FDQuery_EPG_CUR.FieldByName('t_start').AsDateTime;
   t_stop :=DataModule1.FDQuery_EPG_CUR.FieldByName('t_stop').AsDateTime;
   seychas := Now;
   DecodeTime((t_start- seychas),myHour, myMin, mySec, myMilli);
   //p := Trunc((DateTimeToUnix(t_stop) -DateTimeToUnix(t_start))/60) ;
   prodolg := MinutesBetween(T_start,T_stop);
   s_nach := t_start- Now;
   if prodolg =0 then prodolg := 1  ;
   a:= prodolg;
   b:= myMin;
   c:=100;
   //!!!!!!!!!!!
   with DBGrid9.Canvas do
      begin
      //draw rectangle1
      Brush.Color := clWindow ;
      Pen.Color := clRed;
      Rectangle(Rect.Left+1,Rect.Top+2,Rect.Left+Column.Width+1,Rect.Top+20);
      //Rectangle(Rect.Left+1,Rect.Top+2,Rect.Left+Column.Width+1,Rect.Top+20);
      //draw rectangle2
      p :=Trunc( (myMin*100) /prodolg );
       x:= (b*c)/a ;
      rect2_width := p;
      //p := DataModule1.FDQuery_EPG_CUR.FieldByName('t_start_stop').AsInteger;
      Brush.Color := clLime;
      Rectangle(Rect.Left+2,Rect.Top+2,Rect.Left+rect2_width,Rect.Top+20 );

      //draw text
      Brush.Style := bsClear;
      s := IntToStr(p)+'%';
      TextOut((Rect.Left+Column.Width div 2)-TextWidth(s) div 2,Rect.Top+2,s);

      end;

  end;
end;

procedure TForm1.DBGrid9TitleClick(Column: TColumn);
var
  s: string;
begin
  s:=Column.FieldName;
  DataModule1.FDQuery_EPG_CUR.Close;

  if sortFlag then
    DataModule1.FDQuery_EPG_CUR.SQL.Text := 'SELECT * FROM EPG_CUR ORDER BY ' + s + ' DESC'
  else
    DataModule1.FDQuery_EPG_CUR.SQL.Text := 'SELECT * FROM EPG_CUR ORDER BY ' + s + ' ASC';

  sortFlag := not sortFlag;

  DataModule1.FDQuery_EPG_CUR.Open;
end;

procedure TForm1.EditFor10Click(Sender: TObject);
begin
if Form1.Timer1.Enabled = true then
  begin
  Form1.Timer1.Enabled := false;
  Form1.Label39.Visible :=true;
  end
else
  begin
  Form1.Timer1.Enabled := true;
  Form1.Label39.Visible :=false;
  end;

end;

procedure TForm1.Edit16Change(Sender: TObject);
Var i,j:integer;
begin
//
{
For i:=0 To StringGrid1.ColCount-1 Do
   For j:=0 To StringGrid1.RowCount-1 Do
    if pos(AnsiUpperCase(Edit16.Text), AnsiUpperCase(StringGrid1.Cells[i,j]))>0 then
     begin
      StringGrid1.Row:=j;
      StringGrid1.Col:=i;
      Exit;
     end;
}
end;

procedure TForm1.Edit17Change(Sender: TObject);
var
 Text_sql : string;
begin

if Edit17.Text <>'' then
 begin
  Text_sql := 'SELECT * FROM EPG_CUR WHERE channel_en LIKE ';
  Text_sql := Text_sql +chr(39)
    +AnsiUpperCase(Edit17.Text) + '%'+chr(39);
  DataModule1.FDQuery_EPG_CUR.Active := false;
  DataModule1.FDQuery_EPG_CUR.SQL.Clear;
  DataModule1.FDQuery_EPG_CUR.SQL.Add(Text_sql);
  DataModule1.FDQuery_EPG_CUR.Active := true;
 end;

{
if not DataModule1.DataSource_EPG_CUR.DataSet.Locate('channel_ru',Form1.Edit17.Text,[loCaseInsensitive, loPartialKey]) then
    begin
      ShowMessage('Текстовая запись не найдена');
      Form1.Edit17.Clear;
    end;
 }
end;

procedure TForm1.Edit17Click(Sender: TObject);
var
Text_sql   : string;
begin
Text_sql := 'select MAX("start") as max from EGP_program ';

DataModule1.FDQuery_EPG.Active := false;
DataModule1.FDQuery_EPG.SQL.Clear;
DataModule1.FDQuery_EPG.SQL.Add(Text_sql);
DataModule1.FDQuery_EPG.Active := true;

Form1.Edit17.Text :=  DataModule1.FDQuery_EPG.FieldByName('max').AsString;
end;

procedure TForm1.Edit1Change(Sender: TObject);
begin
DataModule1.DataSource_osn.DataSet.Locate('Data',Form1.Edit1.Text,[loCaseInsensitive, loPartialKey]);
end;

procedure TForm1.Edit20Click(Sender: TObject);
begin
if Form1.OpenDialog1.Execute then
  Begin
   Form1.Edit20.Text := Form1.OpenDialog1.FileName ;
   Database_path := Form1.Edit20.Text;
  End;
end;

procedure TForm1.EXCEL1Click(Sender: TObject);
var
  FExcel, FSheet: Variant;
  Rows,  i: Integer;
begin
//DBGrid1  - оплата интернета
if PopupMenu1.PopupComponent = DBGrid1 then
  begin
  CopyGRDToClb(DBGrid1);
 if CheckExcelInstall =true then
    begin
    FExcel := CreateOleObject('Excel.Application');
    FExcel.DisplayAlerts := false;
    FExcel.Visible := true;
    FExcel.Application.EnableEvents := false;
    FExcel.WindowState := -4137;


    FExcel.WorkBooks.Add;
     FSheet := FExcel.Workbooks[1].WorkSheets[1];
    // FExcel.Workbooks[1].WorkSheets[1].Rows[1].HorizontalAlignment := 3;
     //FSheet.Cells[1,2].Columns[1].ColumnWidth := 15;
     //FSheet.Cells[1,2].Columns[1].HorizontalAlignment := 1;
     //FSheet.Cells[I,7].HorizontalAlignment := 3;

    FSheet := FExcel.Workbooks[1].WorkSheets[1]._PasteSpecial;
     Rows := FExcel.ActiveSheet.UsedRange.Rows.Count;
//     Columns := FExcel.ActiveSheet.UsedRange.Columns.Count;
//устанавливаем нужную ширину ячеек
    FExcel.WorkBooks[1].WorkSheets[1].Columns[1].ColumnWidth := 4;
    FExcel.WorkBooks[1].WorkSheets[1].Columns[2].ColumnWidth := 15;
    FExcel.WorkBooks[1].WorkSheets[1].Columns[3].ColumnWidth := 15;
//выравниваем значения в ячейках
   for i := 1  to Rows do
     begin
       FExcel.Workbooks[1].WorkSheets[1].Rows[i].HorizontalAlignment := 3;
     end;
//меняем шрифты
   for i := 1  to Rows do
     begin
      FExcel.WorkBooks[1].WorkSheets[1].Rows[i].Font.Bold := True;
      FExcel.WorkBooks[1].WorkSheets[1].Rows[i].Font.Color := clBlack;
      FExcel.WorkBooks[1].WorkSheets[1].Rows[i].Font.Size := 14;
      FExcel.WorkBooks[1].WorkSheets[1].Rows[i].Font.Name := 'Times New Roman';
     end;

  // обрисовка диапазона ячеек
  for i := 1  to Rows do
  begin
  FExcel.WorkBooks[1].WorkSheets[1].Range['A1:' + chr(ord('C') + 0) +
    IntToStr(i)].Borders.LineStyle := 1;
  FExcel.WorkBooks[1].WorkSheets[1].Range['A1:' + chr(ord('C') + 0) +
    IntToStr(i)].Borders.Weight := 2;
  FExcel.WorkBooks[1].WorkSheets[1].Range['A1:' + chr(ord('C') + 0) +
    IntToStr(i)].Borders.ColorIndex := 1;
  end;

    //FExcel.WorkBooks[1].WorkSheets[1].Rows[2].HorizontalAlignment := 3;
    //FExcel.Workbooks[1].WorkSheets[1].Columns[1].ColumnWidth := 5;
    //FExcel.Workbooks[1].WorkSheets[1].Rows[2].HorizontalAlignment := 3;
    //FSheet.Cells[1,2].Columns[1].HorizontalAlignment := 1;

    end;
  end;

//DBGrid2  - основная таблица
if PopupMenu1.PopupComponent = DBGrid2 then
  begin
  CopyGRDToClb(DBGrid2);
 if CheckExcelInstall =true then
    begin
    FExcel := CreateOleObject('Excel.Application');
    FExcel.DisplayAlerts := false;
    FExcel.Visible := true;
    FExcel.Application.EnableEvents := false;
    FExcel.WindowState := -4137;


    FExcel.WorkBooks.Add;
    FSheet := FExcel.Workbooks[1].WorkSheets[1]._PasteSpecial;
    end;
  end;
//DBGrid3  - таилица выборки
if PopupMenu1.PopupComponent = DBGrid3 then
  begin
  CopyGRDToClb(DBGrid3);
 if CheckExcelInstall =true then
    begin
    FExcel := CreateOleObject('Excel.Application');
    FExcel.DisplayAlerts := false;
    FExcel.Visible := true;
    FExcel.Application.EnableEvents := false;
    FExcel.WindowState := -4137;


    FExcel.WorkBooks.Add;
    FSheet := FExcel.Workbooks[1].WorkSheets[1]._PasteSpecial;
    end;
  end;



//DBGrid4 - таблица электроэнергии
if PopupMenu1.PopupComponent = DBGrid4 then
    begin
  CopyGRDToClb(DBGrid4);
 if CheckExcelInstall =true then
    begin
    FExcel := CreateOleObject('Excel.Application');
    FExcel.DisplayAlerts := false;
    FExcel.Visible := true;
    FExcel.Application.EnableEvents := false;
    FExcel.WindowState := -4137;


    FExcel.WorkBooks.Add;
    FSheet := FExcel.Workbooks[1].WorkSheets[1]._PasteSpecial;
    end;
   end;

//DBGrid5  - таблица оплаты коммнулки
if PopupMenu1.PopupComponent = DBGrid5 then
    begin
  CopyGRDToClb(DBGrid5);
 if CheckExcelInstall =true then
    begin
    FExcel := CreateOleObject('Excel.Application');
    FExcel.DisplayAlerts := false;
    FExcel.Visible := true;
    FExcel.Application.EnableEvents := false;
    FExcel.WindowState := -4137;


    FExcel.WorkBooks.Add;
    FSheet := FExcel.Workbooks[1].WorkSheets[1]._PasteSpecial;
    end;
   end;
//DBGrid6 - таблица моего стажа
if PopupMenu1.PopupComponent = DBGrid6 then
    begin
  CopyGRDToClb(DBGrid6);
 if CheckExcelInstall =true then
    begin
    FExcel := CreateOleObject('Excel.Application');
    FExcel.DisplayAlerts := false;
    FExcel.Visible := true;
    FExcel.Application.EnableEvents := false;
    FExcel.WindowState := -4137;


    FExcel.WorkBooks.Add;
    FSheet := FExcel.Workbooks[1].WorkSheets[1]._PasteSpecial;
    end;
   end;

//DBGrid7 - таблица паролей
if PopupMenu1.PopupComponent = DBGrid7 then
    begin
  CopyGRDToClb(DBGrid7);
 if CheckExcelInstall =true then
    begin
    FExcel := CreateOleObject('Excel.Application');
    FExcel.DisplayAlerts := false;
    FExcel.Visible := true;
    FExcel.Application.EnableEvents := false;
    FExcel.WindowState := -4137;


    FExcel.WorkBooks.Add;
    FSheet := FExcel.Workbooks[1].WorkSheets[1]._PasteSpecial;
    end;
   end;


// GetActiveOleObject;
//if  GetActiveOleObject = true then
//if CheckExcelRun = true then


end;

procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
var
//Что касается резервного копирования
 arch_fie_name : string;
 path_sohran_bak :string;
// kol_files :integer;
 S  : string;
 s0 : string;
 s1: string;
 s2: string;
 i : integer;

 Path,Path_1,  Mask : String;
 Sr         : TSearchRec;
//переменные для ротации
date_file_rotac : string;
Mask_rotac      : string;
File_rotac      : string;
//что касается проецесса
si : TStartupInfo;
pi : TProcessInformation;
cmdLine : pChar;
app : pChar;
//для архивации
AppToRun :string; //путь до winrar
filename:string; //имя файла
date_file :string;
CommandLine:string; //строка параметров
Rlst: LongBool; //результат выполнения
StartUpInfo: TStartUpInfo; //параметры будущего процесса
ProcessInfo: TProcessInformation; //Отслеживание выполнения
Error:integer; //номер ошибкок
ExitCode: Cardinal; //код завершения
path_arch :string;
Work_dir  :string;
Copy_file  :string;

begin
  if Form1.Backup_drive = '' then
  begin
  // Form1.ComboBoxEx4.Text := 'D:\';
  ShowMessage('Ты не установил диск для архивов перейди на вкладку Set, и установи все нужные значения.');
  Form1.PageControl1.ActivePage := TabSheet10;
  Action:=caNone;
  end;

Path:= ExtractFilePath( Application.ExeName );
Path_1 := Path + 'Base_Backup\';
//Определение существования дирректории для BackUp
     if System.SysUtils.DirectoryExists(Path_1) then
       begin
         path_sohran_bak := Path_1;
       end
     else
       begin
         CreateDir(Path_1);
         path_sohran_bak := Path_1;
       end;
//============================================
SetCurrentDir(Path_1);
if FindFirst('*', faAnyFile, Sr) = 0 then
begin
repeat
if FileDateToDateTime(Sr.time) <= (date-3) then
begin
deletefile(Sr.name);
end;
until FindNext(Sr) <> 0;
FindClose(Sr);
end;
//Поиск файла по маске
   Mask :='*_'+Copy(today_str,1,2)+'_'
   +Copy(today_str,4,2)+'_'
   +Copy(today_str,7,4)+ '*';
if FindFirst(Path_1+Mask, faAnyFile, sr) = 0 then
  begin
  // Должен освободить ресурсы, используемые этими успешными, поисками
    DeleteFile(Path_1+sr.Name);
  end;
  FindClose(sr);
//перемещаем файл базы
//Если есть дирректория - удаляем
  if DirectoryExists('D:\nd') then DeleteDir('D:\nd');
//Если есть дирректория - удаляем
  if DirectoryExists('D:\nd') then DeleteDir('D:\nd');
//Создаем заново
Path_1 := 'D:\nd\';
     if System.SysUtils.DirectoryExists(Path_1) then
       begin
         path_arch := Path_1;
       end
     else
       begin
         CreateDir(Path_1);
         path_arch := Path_1;
       end;

  Work_dir := ExtractFileDir(Application.ExeName);
  Work_dir := StringReplace(Work_dir, '\', '/', [rfReplaceAll]);
  Copy_file := Work_dir +'/Bill_New.db3';
  CopyFile(PChar(Copy_file), 'd:/nd/Bill_New.db3',true);

//Определяем путь к WinRAR.
AppToRun :='C:\Program Files\WinRAR\rar.exe';

 filename:='d:\nd\Bill_New.db3';
   date_file := FormatDateTime('dd_mm_yyyy"-"hh_nn_ss', Now);

Path_1 := ExtractFileDir(Application.ExeName)+'\Base_Backup\';
CommandLine:='"'+AppToRun+'"' +'  a '
 +Path_1 +'Bill_base_'+ date_file + '.rar '+filename;
//arch_fie_name := FormatDateTime('dd_mm_yyyy"-"hh_nn_ss', Now);
 //Заполнение нулями всего StartUpInfo.
FillChar(StartUpInfo, SizeOf(TStartUpInfo), 0);
//После этого выставляем в нем некоторые параметры.
 with StartUpInfo do
begin

//Содержит количество байтов, занимаемых структурой TStartUpInfo.
//Обязательно для заполнения. Инициализируйте как SizeOf(TStartUpInfo).
cb := SizeOf(TStartUpInfo);
//Содержит набор флагов, позволяющих управлять созданием дочернего процесса.
//Показываем окно, курсор - часики.
dwFlags := STARTF_USESHOWWINDOW or STARTF_FORCEONFEEDBACK;
//Определяет как должно выглядеть окно запущенного приложения.
//Нормальное отображение
wShowWindow := SW_SHOWNORMAL;
end;
//Сам запуск WinRAR с параметрами.
//По сути, мы запускаем следующее:
//WinRar.exe a -ep "<путь до архива>" "<путь до файла>"
//showmessage(AppToRun+CommandLine);
//my:= AppToRun+CommandLine;
Rlst:= CreateProcess(PChar(AppToRun), PChar(CommandLine),
nil, nil, false, NORMAL_PRIORITY_CLASS,
nil, nil, StartUpInfo, ProcessInfo);

//Отслеживаем выполнение.
//Если запуск успешен
if Rlst then
with ProcessInfo do
begin
//Ждем завершения инициализации.
WaitForInputIdle(hProcess, INFINITE);
//Ждем завершения процесса.
WaitforSingleObject(ProcessInfo.hProcess, INFINITE);
//Получаем код завершения.
GetExitCodeProcess(ProcessInfo.hProcess, ExitCode);
//Закрываем дескриптор процесса.
CloseHandle(hThread);
//Закрываем дескриптор потока.
CloseHandle(hProcess);
end
//Иначе
else
  begin
//В случае ошибки - выводим на экран сообщение.
Error := GetLastError;
MessageDlg(SysErrorMessage(Error), mtError, [mbOk], 0);
  end;
DeleteDir('D:\nd');
end;

procedure TForm1.FormCreate(Sender: TObject);
var
  Ini: Tinifile;
  w: dword;
  Root: string;
  i,DiskType: integer;
  f:textfile; // Dalee textovie peremennie
  Present: TDateTime;      // текущая дата и время
  Year, Month, Day : Word; // год, месяц и число, как


begin
//Read ini file
  {диск А это привод гибких дисков,если он не нужен то
  цикл до 24 и
  Root[1] := Char(Ord('B') + i); }
  w := GetLogicalDrives;
  Root := '#:\';
  for i := 0 to 25 do
 begin
    Root[1] := Char(Ord('A') + i);
    if (W and (1 shl i)) > 0 then
      if (GetDriveType(Pchar(Root)) = DRIVE_FIXED)  then
  begin
   Form1.ComboBoxEx4.Items.Add(Root);
  end;
  Form1.DateTimePicker5.Date := Now;
  Form1.Edit21.Text := '';

 end;

//открываем файл  /////////////////////
  Ini:=TiniFile.Create(extractfilepath(paramstr(0))+'My_Ini.ini');
  Backup_drive := Ini.ReadString('Main','Drive_Backup',Form1.ComboBoxEx1.Text);
  Database_path := Ini.ReadString('Main','Database_path',Form1.ComboBoxEx1.Text);
  ejem_platej   := Ini.ReadInteger('Main','Ежемесячный платёж',Form1.SpinEdit4.Value);

  Form1.SpinEdit1.Value := ejem_platej;
  Form1.SpinEdit4.Value := ejem_platej;

  Form1.Edit20.Text := Database_path;
 // Form1.Width:=Ini.ReadInteger('Size','Width',100);
  //последнее значение (100) это значение по умолчанию (default)
 // Form1.Height:=Ini.ReadInteger('Size','Height',100);
 // Form1.Left:=Ini.ReadInteger('Position','X',10);
 // Form1.Top:=Ini.ReadInteger('Position','Y',10);
  Ini.Free;

//=====================================
 tm_work := 1200 ;
 Form1.Edit10.Color := clGreen;
 Form1.Edit10.Text := IntToStr(tm_work);
 //
 today_str := DateToStr(Now);
 //
Form1.DBGrid1.PopupMenu := PopupMenu1;
Form1.DBGrid2.PopupMenu := PopupMenu1;
Form1.DBGrid3.PopupMenu := PopupMenu1;
Form1.DBGrid4.PopupMenu := PopupMenu1;
Form1.DBGrid5.PopupMenu := PopupMenu1;
Form1.DBGrid6.PopupMenu := PopupMenu1;
Form1.DBGrid7.PopupMenu := PopupMenu1;

Form1.Edit10.Color := RGB(102, 204, 240);  ;
Form1.PageControl1.ActivePage := TabSheet1;

    GetKeyboardLayoutName(Layout);
    if Layout='00000409' then
    Label48.Caption:= NameKeyboardLayout(GetActiveKbdLayout);
    if Layout='00000419' then
    Label48.Caption:= NameKeyboardLayout(GetActiveKbdLayout);

  //  LoadKeyboardLayout(StrCopy(Layout,'00000409'),KLF_ACTIVATE);

  {активная раскладка в своей программе}
 //  Label48.Caption:= NameKeyboardLayout(GetActiveKbdLayout);
   // Label48.Caption:= NameKeyboardLayout(GetActiveKbdLayout);
//=========================
 Present:= Now; // получить текущую дату
 DecodeDate(Present, Year, Month, Day);
 //Label2.Caption := 'Сегодня '+IntToStr(Day)+ ' ' +
//  stMonth[Month] + ' '+ IntToStr(Year)+
// ' года, '+ stDay[DayOfWeek(Present)];

 // настроить и запустить таймер
 Timer3.Interval := 1000; // период сигналов таймера 1 с
 Timer3.Enabled := True; // пуск таймера

//=========================


end;

procedure TForm1.FormDestroy(Sender: TObject);
var
  Ini: Tinifile; //необходимо создать объект, чтоб потом с ним работатьbegin
begin
{
//создали файл в директории программы
  Ini:=TiniFile.Create(extractfilepath(paramstr(0))+'My_Ini.ini');
  Ini.WriteString('Main','Drive_Backup',Form1.Backup_drive);
  Ini.WriteInteger('Size','Width',form1.width);
  Ini.WriteInteger('Size','Height',form1.height);
  Ini.WriteInteger('Position','X',form1.left);
  Ini.WriteInteger('Position','Y',form1.top);
  Ini.Free;
 }
end;

procedure TForm1.FormShow(Sender: TObject);
var
 Text_SQL        : string;
 konec_mes       : string;
 konec_mes_lab2  : string;
 mes_days_str    : string;
 mes_days        : integer;
 Ostatok_na_konec: integer;

 Trawm            : integer;
 Zapl_inet        : integer;
 Gips             : integer;
 Sn_Gips          : integer;
 Den_Gizni        : integer;
// Sobitie          : string;

 MAX_Data_opl_str : string;
 MAX_Data_opl     : TDate;

 like             : string;
 plat             : string;
 ejed_day_pay     : real;
 i                : real;

 god_tek          : integer;
 mes_tek          : integer;

 Summa_ostatok        : real;
 Ostatok_na_konec_tek : real;
 chet_dney            : integer;

 DaysInAMonth_tek     : integer;
 ejed_day_pay_tek     : real;

 Date1                : TDateTime;
 Date2                : TDateTime;
 nach_data            : string;
 kon_data             : string;

 for_cicle            : integer;
// record_count         : integer;
 zam                  : string;
 s1                   : string;
 dat_zam              : string;
 do_dny_ost           : integer;

// flag_edit9           : integer;
 i_nap              : integer;
 fileName           : string;
 fileDate           : Integer;

 Y, W, D: Word;
label
  GotoLabel;
begin

//FDConnection_osn
DataModule1.FDConnection_osn.DriverName :='SQLite'  ;
DataModule1.FDConnection_osn.Params.Database :=Database_path ;
//ExtractFileDir(Application.ExeName)+'\'+ 'Bill_New.db3';
DataModule1.FDConnection_osn.Connected := true ;


/////// поднимаем FDQeuery коммуналки
   Text_SQL := 'Select * From Svet';
   DataModule1.FDQuery_Svet.Active := false;
   DataModule1.FDQuery_Svet.SQL.Clear;
   DataModule1.FDQuery_Svet.SQL.Add(Text_SQL);
   DataModule1.FDQuery_Svet.Active := true;
   DataModule1.FDQuery_Svet.Last;

   Text_SQL := 'Select * From Kvartira';
   DataModule1.FDQuery_Kvartira.Active := false;
   DataModule1.FDQuery_Kvartira.SQL.Clear;
   DataModule1.FDQuery_Kvartira.SQL.Add(Text_SQL);
   DataModule1.FDQuery_Kvartira.Active := true;
   DataModule1.FDQuery_Kvartira.Last;

   DataModule1.FDQuery_Stage.Active := true;
   //DataModule1.FDQuery_Stage.Last;

   DataModule1.FDQuery_Log.Active := true;
   DataModule1.FDQuery_Log.First;

//=============================================
// Наполняем ComboBoxEx1
 Text_Sql := 'select * from EGP_Loc';
DataModule1.FDQuery_C3.Active :=false;
DataModule1.FDQuery_C3.SQL.Clear;
DataModule1.FDQuery_C3.SQL.Add(Text_SQL);
DataModule1.FDQuery_C3.Active :=true;

DataModule1.FDQuery_C3.Last;
DataModule1.FDQuery_C3.First;
for i_nap := 0  to DataModule1.FDQuery_C3.RecordCount -1  do
begin
s1 := DataModule1.FDQuery_C3.FieldByName('Name_chanel').AsString;
Form1.ComboBoxEx1.Items.Add(DataModule1.FDQuery_C3.FieldByName('Name_chanel').AsString);

DataModule1.FDQuery_C3.Next;
end;
//======================================================
// Наполняем ComboBoxEx2
 Text_Sql := 'select * from EGP_Loc';
DataModule1.FDQuery_C3.Active :=false;
DataModule1.FDQuery_C3.SQL.Clear;
DataModule1.FDQuery_C3.SQL.Add(Text_SQL);
DataModule1.FDQuery_C3.Active :=true;

DataModule1.FDQuery_C3.Last;
DataModule1.FDQuery_C3.First;
for i_nap := 0  to DataModule1.FDQuery_C3.RecordCount -1  do
begin
s1 := DataModule1.FDQuery_C3.FieldByName('Name_chanel').AsString;
Form1.ComboBoxEx2.Items.Add(DataModule1.FDQuery_C3.FieldByName('Name_chanel').AsString);

DataModule1.FDQuery_C3.Next;
end;
//=======================================================
// Наполняем ComboBoxEx3
 Text_Sql := 'select * from EGP_channel';
DataModule1.FDQuery_C3.Active :=false;
DataModule1.FDQuery_C3.SQL.Clear;
DataModule1.FDQuery_C3.SQL.Add(Text_SQL);
DataModule1.FDQuery_C3.Active :=true;

DataModule1.FDQuery_C3.Last;
DataModule1.FDQuery_C3.First;
for i_nap := 0  to DataModule1.FDQuery_C3.RecordCount -1  do
begin
s1 := DataModule1.FDQuery_C3.FieldByName('display_name').AsString;
//Form1.ComboBoxEx3.Items.Add(DataModule1.FDQuery_C3.FieldByName('display_name').AsString);
Form1.ComboBoxEx3.Items.Add(s1);
DataModule1.FDQuery_C3.Next;
end;
//=======================================================
// Наполняем ComboBox5 (Восход солнца)
Form1.ComboBox5.Items.Add('world-weather.ru');
//https://world-weather.ru/pogoda/russia/donetsk/sunrise/?ysclid=lb4qg3rfwe370408491
Form1.ComboBox5.Items.Add('ru.meteocast.net');
//https://ru.meteocast.net/sunrise-sunset/ua/donetsk/
Form1.ComboBox5.Items.Add('time-in.ru');
//https://time-in.ru/sun/donetsk?ysclid=lb4rlrbdvi595900714
Form1.ComboBox5.Items.Add('voshod-solnca.ru');
//https://voshod-solnca.ru/equinoxes/%D0%B4%D0%BE%D0%BD%D0%B5%D1%86%D0%BA_(%D1%83%D0%BA%D1%80%D0%B0%D0%B8%D0%BD%D0%B0)/NaN


//-----------------------------------------------------------
// Наполняем ComboBox3 (Погода)
Form1.ComboBox3.Items.Add('gismeteo.ru');
Form1.ComboBox3.Items.Add('yandex.ru');
Form1.ComboBox3.Items.Add('pogoda7.ru');
Form1.ComboBox3.Items.Add('pogoda.mail.ru');
//Form1.ComboBox3.Items.Add('rp5.ru');
//-----------------------------------------------------------
// Наполняем ComboBox7 (Именины сегодня)
Form1.ComboBox7.Items.Add('my-calend.ru');
Form1.ComboBox7.Items.Add('nashiimena.ru');
Form1.ComboBox7.Items.Add('www.calend.ru');
//===========================================================
// Наполняем ComboBox8 (Какой сегодня праздник)
Form1.ComboBox8.Items.Add('kakoysegodnyaprazdnik.ru');
Form1.ComboBox8.Items.Add('calendar.online.ua');

ejem_platej :=Form1.SpinEdit1.Value;

day:= StrToInt(copy(today_str,1,2));
mes:= StrToInt(copy(today_str,4,2));
god:= StrToInt(copy(today_str,7,4));

Form1.Label1.Caption:='Ежемесячный платёж';
Form1.SpinEdit1.Alignment := taRightJustify;

Text_label2 :='Остаток на счете на ';

Form1.SpinEdit2.Alignment := taRightJustify;
Form1.Label3.Caption:= today_str;

Text_SQL := 'SELECT date(:p1,:p2,:p4) as Date;';
DataModule1.FDQuery_old_1.Active :=false;
DataModule1.FDQuery_old_1.SQL.Clear;
DataModule1.FDQuery_old_1.SQL.Add(Text_SQL);
DataModule1.FDQuery_old_1.Params.ParamByName('p1').Value :='now' ;
DataModule1.FDQuery_old_1.Params.ParamByName('p2').Value :='start of month' ;
DataModule1.FDQuery_old_1.Params.ParamByName('p4').Value := '-1 day';
DataModule1.FDQuery_old_1.Active :=true;
//konec_mes :=DataModule1.FDQuery_old_1.FieldByName('Date').AsString;
//Для МАТРИЦЫ Дата списания
konec_mes :=DataModule1.FDQuery_old_1.FieldByName('Date').AsString;

konec_mes_lab2 :=Copy(konec_mes,9,2);
konec_mes_lab2 :=konec_mes_lab2+'.' + Copy(konec_mes,6,2);
konec_mes_lab2 :=konec_mes_lab2+'.' + Copy(konec_mes,1,4);
Form1.Label2.Caption := Text_label2 + konec_mes_lab2;

Text_SQL := 'SELECT date(:p1,:p2,:p3,:p4) as Date;';
DataModule1.FDQuery_old_1.Active :=false;
DataModule1.FDQuery_old_1.SQL.Clear;
DataModule1.FDQuery_old_1.SQL.Add(Text_SQL);
DataModule1.FDQuery_old_1.Params.ParamByName('p1').Value :='now' ;
DataModule1.FDQuery_old_1.Params.ParamByName('p2').Value :='start of month' ;
DataModule1.FDQuery_old_1.Params.ParamByName('p3').Value := '+1 month';
DataModule1.FDQuery_old_1.Params.ParamByName('p4').Value := '-1 day';
DataModule1.FDQuery_old_1.Active :=true;

mes_days_str :=DataModule1.FDQuery_old_1.FieldByName('Date').AsString;
mes_days :=StrToInt(copy(mes_days_str,9,2));

Text_SQL := 'select * from Bill_New  where Data =:p1';
DataModule1.FDQuery_old_1.Active :=false;
DataModule1.FDQuery_old_1.SQL.Clear;
DataModule1.FDQuery_old_1.SQL.Add(Text_SQL);
DataModule1.FDQuery_old_1.Params.ParamByName('p1').Value := konec_mes;
DataModule1.FDQuery_old_1.Active :=true;      // ТУТ НА 32 ОШИБКА

Ostatok_na_konec := DataModule1.FDQuery_old_1.FieldByName('Ostatok').AsInteger;

//вычисляем Trawm
Trawm := DaysBetween(StrToDate(today_str), StrToDate(Trawm_ish)  ) ;
Form1.Label7.Caption := Form1.Label7.Caption + '  ' + IntToStr(Trawm);
Form1.Edit2.Text :=  IntToStr(Trawm);

//вычисляем Sn_Gips

Sn_Gips :=  DaysBetween(StrToDate(today_str), StrToDate(Sn_Gips_ish)  ) ;
Form1.Label8.Caption := Form1.Label8.Caption + '  ' + IntToStr(Sn_Gips);
Form1.Edit4.Text :=  IntToStr(Sn_Gips);

//вычисляем Gips
Gips := DaysBetween(StrToDate(today_str), StrToDate(Gips_ish)  ) ;
Form1.Edit3.Text :=  IntToStr(Gips);

//вычисляем Zapl_inet
Text_SQL := 'select MAX(Data_opl) from Opl_table ';
DataModule1.FDQuery_old_1.Active :=false;
DataModule1.FDQuery_old_1.SQL.Clear;
DataModule1.FDQuery_old_1.SQL.Add(Text_SQL);
DataModule1.FDQuery_old_1.Active :=true;

MAX_Data_opl_str := DataModule1.FDQuery_old_1.FieldByName('MAX(Data_opl)').AsString;
MAX_Data_opl_str := Copy(MAX_Data_opl_str,9,2)
     +'.'+Copy(MAX_Data_opl_str,6,2)
     +'.'+Copy(MAX_Data_opl_str,1,4);

MAX_Data_opl := StrToDate(MAX_Data_opl_str) ;
Zapl_inet := DaysBetween(StrToDate(today_str),
             MAX_Data_opl) ;

Form1.Edit5.Text := IntToStr(Zapl_inet);
//День жизни
//Den_Gizni :=  DaysBetween(StrToDate('09.05.2021'), StrToDate('18.06.1967')  ) ;
Den_Gizni :=  DaysBetween(StrToDate(today_str), StrToDate('18.06.1967')  ) ;
Form1.Edit7.Text := IntToStr(Den_Gizni);
//================================================================
Form1.SpinEdit2.Value := Ostatok_na_konec;
//like :=IntToStr(god)+'-'+ IntToStr(mes)+'%';
like := Copy(today_str,7,4) + '-'+Copy(today_str,4,2)+ '%';
//like := like + '-'+Copy(today_str,4,2);
Text_SQL := 'select sum(Sum_opl) as Sum_opl from Opl_table where 0=0 and  Data_opl like :p1';
DataModule1.FDQuery_old_1.Active :=false;
DataModule1.FDQuery_old_1.SQL.Clear;
DataModule1.FDQuery_old_1.SQL.Add(Text_SQL);
DataModule1.FDQuery_old_1.Params.ParamByName('p1').Value := like;
DataModule1.FDQuery_old_1.Active :=true;

plat := DataModule1.FDQuery_old_1.FieldByName('Sum_opl').AsString;
if plat = '' then plat :='0' ;


ejed_day_pay := ejem_platej /mes_days;
i := (Ostatok_na_konec+ StrToInt(plat))   - (day * ejed_day_pay);
Form1.SpinEdit3.Value :=Round(i);

//=========================================================
//Оплата начальна инициализация
Text_SQL := 'SELECT * FROM Opl_table';
DataModule1.FDQuery_opl.Active :=false;
DataModule1.FDQuery_opl.SQL.Clear;
DataModule1.FDQuery_opl.SQL.Add(Text_SQL);
DataModule1.FDQuery_opl.Active :=true;
DataModule1.FDQuery_opl.Last;

   Form1.DBGrid1.Columns[0].Width := 25;
   Form1.DBGrid1.Columns[0].Alignment := taCenter;
   Form1.DBGrid1.Columns[1].Width := 90;
   Form1.DBGrid1.Columns[1].Alignment := taCenter;
   Form1.DBGrid1.Columns[2].Width := 70;
   Form1.DBGrid1.Columns[2].Alignment := taCenter;
   Form1.DBGrid1.Columns[3].Width := 150;
   Form1.DBGrid1.Columns[3].Alignment := taCenter;

Text_SQL :='select * from Bill_New ';
DataModule1.FDQuery_osn.Active :=false;
DataModule1.FDQuery_osn.SQL.Clear;
DataModule1.FDQuery_osn.SQL.Add(Text_SQL);
DataModule1.FDQuery_osn.Active :=true;

   Form1.DBGrid2.Columns[0].Width := 50;
   Form1.DBGrid2.Columns[0].Alignment:=taCenter;
   Form1.DBGrid2.Columns[1].Width := 50;
   Form1.DBGrid2.Columns[1].Alignment:=taCenter;
   Form1.DBGrid2.Columns[2].Width := 105;
   Form1.DBGrid2.Columns[3].Width := 80;
   Form1.DBGrid2.Columns[3].Alignment:=taCenter;
   Form1.DBGrid2.Columns[4].Width := 50;
   Form1.DBGrid2.Columns[4].Alignment:=taCenter;
//   Form1.DBGrid2.Columns[4].Font
   Form1.DBGrid2.Columns[5].Width := 80;
   Form1.DBGrid2.Columns[6].Width := 80;
   Form1.DBGrid2.Columns[7].Width := 80;
   Form1.DBGrid2.Columns[8].Width := 80;
   Form1.DBGrid2.Columns[9].Width := 80;
   Form1.DBGrid2.Columns[10].Width := 210;

   Form1.Edit1.Text := today_str;
   Form1.ComboBox2.Text := DataModule1.FDQuery_osn.FieldByName('Sobitie').AsString;
   if DataModule1.FDQuery_osn.FieldByName('Pil').AsString <>''then
   begin
    Form1.ComboBox1.Text := DataModule1.FDQuery_osn.FieldByName('Pil').AsString;
   end
  else  Form1.ComboBox1.Text := 'Н/З ' ;

//============================================================================
god_tek := god;
mes_tek := mes;

Summa_ostatok := Form1.SpinEdit3.Value;
Ostatok_na_konec_tek := Summa_ostatok;

//Ostatok_na_konec;
chet_dney := 0;
//идем по месяцам
//==================================================
 While Ostatok_na_konec_tek > ejem_platej do
  begin
     DaysInAMonth_tek := DaysInAMonth(god_tek, mes_tek);
     ejed_day_pay_tek := ejem_platej / DaysInAMonth_tek;
     Ostatok_na_konec_tek :=Ostatok_na_konec_tek
      - (DaysInAMonth_tek*ejed_day_pay_tek) ;

      if Ostatok_na_konec_tek > 0 then
      begin
       Inc(mes_tek);
       chet_dney := chet_dney + DaysInAMonth_tek;
      end;
  end;
//===================================================
//идем по дням
While Ostatok_na_konec_tek > 0 do
  begin
     DaysInAMonth_tek := DaysInAMonth(god_tek, mes_tek);
     ejed_day_pay_tek := ejem_platej / DaysInAMonth_tek;
     Ostatok_na_konec_tek :=Ostatok_na_konec_tek
      - ejed_day_pay_tek ;
      Inc(chet_dney);
  end;
Form1.Edit6.Text := DateToStr(IncDay(Date, chet_dney));
//Form1.Edit23.Text := IntToStr( chet_dney);
Form1.Edit5.Text :=Form1.Edit5.Text+'   Осталось: '+ IntToStr(chet_dney);
//============================================================================
  Date1:= StrToDate('18.06.1967')  ;
  Date2:= StrToDate(today_str)  ;

  Form1.Edit8.Text := FormatAge(Now - StrToInt(Edit7.Text) , Now, True);

 // Form1.PageControl1.ActivePage := TabSheet1;

//===============================================================
// Вычисляем сколько дней до ближайшего события

Text_sql := 'SELECT * FROM Bill_New  ';
Text_sql :=Text_sql +'  where Data BETWEEN :p1 and :p2 ';

nach_data :=Copy(today_str,7,4)
     +'-'+Copy(today_str,4,2)
     +'-'+Copy(today_str,1,2);



kon_data := Copy(DateToStr(IncYear(Now)),7,4)
            +'-'+'01'
            +'-'+'01' ;
{
     kon_data := Copy(today_str,7,4)
      +'-'+Copy(today_str,4,2)
      +'-'+
      IntToStr(DaysInAMonth(god, mes));
 }
DataModule1.FDQuery_excel.Active := false;
DataModule1.FDQuery_excel.SQL.Clear;
DataModule1.FDQuery_excel.SQL.Add(Text_SQL);

DataModule1.FDQuery_excel.Params.ParamByName('p1').Value := nach_data;
DataModule1.FDQuery_excel.Params.ParamByName('p2').Value := kon_data;
DataModule1.FDQuery_excel.Active := true;

 //record_count :=  DataModule1.FDQuery_excel.RecordCount;
 DataModule1.FDQuery_excel.First;
 for for_cicle := 0 to DataModule1.FDQuery_excel.RecordCount do
 begin
  zam:= DataModule1.FDQuery_excel.FieldByName('Zam').AsString;
  if zam <> ''   then
   begin
    s1 := DataModule1.FDQuery_excel.FieldByName('Zam').AsString;
    dat_zam := DataModule1.FDQuery_excel.FieldByName('Data').AsString;

    Form1.Label29.Caption := Form1.Label29.Caption
     +' '+ zam;

    goto GotoLabel;
   end;
 DataModule1.FDQuery_excel.Next;
 end;

GotoLabel:

if dat_zam <>'' then
  begin
    do_dny_ost :=DaysBetween(StrToDate(dat_zam) ,StrToDate(today_str) );
    if do_dny_ost <= 1 then Form1.Edit9.Color := clRed;
    if do_dny_ost >  0 then Form1.Edit9.Text :=dat_zam+'     Осталось:  '+FormatAge(Now, Now - do_dny_ost);
    if do_dny_ost =  0 then Form1.Edit9.Text :='Сегодня!';

   end
 else
  begin
   Form1.Edit9.Text :='нет событий до конца года.';
  end;
   DataModule1.FDQuery_Svet.Active := true;
   DataModule1.FDQuery_Svet.Last;

   DataModule1.FDQuery_Kvartira.Active := true;
   DataModule1.FDQuery_Kvartira.Last;

   DataModule1.FDQuery_Stage.Active := true;
//===============================
//Здесь была ошибка только 01.03.2023 !!!!!
DataModule1.FDQuery_Stage.Last;
//===============================
DataModule1.DataSource_osn.DataSet.Locate('Data',Form1.Edit1.Text,[loCaseInsensitive, loPartialKey]);
DataModule1.FDQuery_excel.Active := false;

  fileName := 'D:\unicast.m3u';
  fileDate := FileAge(fileName);
   if fileDate > -1 then
   Form1.Edit14.Text :=( 'посл.изм   ' +
     DateToStr(FileDateToDateTime(fileDate)));
   if fileDate > -1 then
   begin
    if DaysBetween(FileDateToDateTime(fileDate),now) > 1 then
     begin
     Form1.Edit14.Color := clRed ;
     end;
   end;


  fileName := 'D:\New_Bill_Tab\Win64\Debug\EGP_cur\xmltv_trk.xml';
  fileDate := FileAge(fileName);
   if fileDate > -1 then
   Form1.Edit15.Text :=( 'посл.изм   ' +
     DateToStr(FileDateToDateTime(fileDate)));

     //fileDate := DaysBetween(FileDateToDateTime(fileDate),now);
 if fileDate > -1 then
 begin
  if DaysBetween(FileDateToDateTime(fileDate),now) > 1 then
    begin
     Form1.Edit15.Color := clRed ;
    end;
 end;

 DecodeDateWeek(Now, Y, W, D); // извлекаем значения из текущей даты и времени
//  ShowMessage('Сейчас ' + IntToStr(Y) + ' год ' + IntToStr(W) + ' неделя ' + IntToStr(D) + ' день');
Form1.Edit18.Text := IntToStr(W) + ' неделя';
Form1.Edit19.Text :=FloatToStr(GetDays(Date)) +' день в году';
end;



procedure TForm1.TabSheet10Show(Sender: TObject);
begin
Form1.ComboBoxEx4.Text := Form1.Backup_drive;
end;

procedure TForm1.TabSheet11Show(Sender: TObject);
var
   Text_Sql : string;
begin
{
   Text_Sql := 'select * from Dengi ';
   DataModule1.FDQuery_Dengi.Active :=false;
   DataModule1.FDQuery_Dengi.SQL.Clear;
   DataModule1.FDQuery_Dengi.SQL.Add(Text_SQL);
   DataModule1.FDQuery_Dengi.Active := true;
   DataModule1.FDQuery_Dengi.Last;
   //Form1.DBGrid_dengi.SetFocus;
 }
  Form1.PageControl2.ActivePage :=TabSheet12;
end;



procedure TForm1.TabSheet12Show(Sender: TObject);
var
  Text_sql    :   string;
begin
//if PageControl2.ActivePage = TabSheet12 then //Zarpl
//  begin
    Form1.Button_Dengi.Caption :=
     'Загрузить выписку по зарплате';
     Form1.Button_Och.Caption :=
     'Очистить выписку по зарплате';

    Text_sql :='SELECT * FROM Vipiska_Zarpl';

    DataModule1.FDQuery_Dengi_Zarpl.Active := false;
    DataModule1.FDQuery_Dengi_Zarpl.SQL.Clear;
    DataModule1.FDQuery_Dengi_Zarpl.SQL.Add(Text_sql);
    DataModule1.FDQuery_Dengi_Zarpl.Active := true;
    DataModule1.FDQuery_Dengi_Zarpl.Last;
    Form1.DBGrid10.SetFocus
//  end;
end;

procedure TForm1.TabSheet13Show(Sender: TObject);
var
  Text_sql    :   string;

begin
// if PageControl2.ActivePage = TabSheet13 then //Pens
//  begin
    Form1.Button_Dengi.Caption :=
     'Загрузить выписку по пенсии';
     Form1.Button_Och.Caption :=
     'Очистить выписку по пенсии';

    Text_sql :='SELECT * FROM Vipiska_Pens';

    DataModule1.FDQuery_Dengi_Pens.Active := false;
    DataModule1.FDQuery_Dengi_Pens.SQL.Clear;
    DataModule1.FDQuery_Dengi_Pens.SQL.Add(Text_sql);
    DataModule1.FDQuery_Dengi_Pens.Active := true;
    DataModule1.FDQuery_Dengi_Pens.Last;

    Form1.DBGrid11.SetFocus


//  end;

end;

procedure TForm1.TabSheet3Show(Sender: TObject);
var
 fdate : string;
begin
fdate :='01.'+ Copy(today_str,4,Length(today_str));
Form1.DateTimePicker1.Date := StrToDate(fdate);
Form1.DateTimePicker2.Date := StrToDate(today_str);
Form1.DBGrid3.SetFocus;
end;

procedure TForm1.TabSheet4Show(Sender: TObject);
begin
//
Form1.ComboBoxEx1.SetFocus;
end;

procedure TForm1.TabSheet6Show(Sender: TObject);
begin
//
DataModule1.FDQuery_Voda.Active := true;
DataModule1.FDQuery_Voda.Last;
end;

procedure TForm1.TabSheet7Show(Sender: TObject);
var
  Stage      : integer;
  From_Date  : Tdate;
  To_Date    : Tdate;
begin
From_Date  := DataModule1.FDQuery_Stage.FieldByName('From_Date').AsDateTime;
To_Date    := DataModule1.FDQuery_Stage.FieldByName('To_Date').AsDateTime;
Stage := DaysBetween(To_Date, From_Date  ) ;
Form1.Edit11.Text := IntToStr(Stage);
Form1.Edit13.Text := FormatAge(Now - StrToInt(Form1.Edit11.Text), Now, True);// 27 л. 7 мес. 24 д.
Form1.Edit12.Text := '';

Form1.DBGrid6.SetFocus;

end;

procedure TForm1.Timer1Timer(Sender: TObject);
begin
//
tm_work :=tm_work -1;

  if tm_work <= 300  then
  begin
   Form1.Edit10.Color := clYellow;
   Form1.Edit10.Font.Color := clBlack;
  end;

  if tm_work <= 100  then
  begin
   Form1.Edit10.Color := clRed;
   Form1.Edit10.Font.Color := clWhite;
  end;

  if tm_work = 0  then Form1.Close;

Form1.Edit10.Text := IntToStr(tm_work);
//Form1.Button34Click(Sender);
end;

procedure TForm1.Timer2Timer(Sender: TObject);
begin
//Form1.Button34Click(Sender);
end;

procedure TForm1.Timer3Timer(Sender: TObject);
begin
ShowTime;
end;

procedure TForm1.Menu11Click(Sender: TObject);
begin
//DBGrid1  - оплата интернета
if PopupMenu1.PopupComponent = DBGrid1 then
  CopyGRDToClb(DBGrid1);
//DBGrid2  - основная таблица
if PopupMenu1.PopupComponent = DBGrid2 then
  CopyGRDToClb(DBGrid2);
//DBGrid2  - редактирование основной таблицы
if PopupMenu1.PopupComponent = Form3.DBGrid2 then
  CopyGRDToClb(Form3.DBGrid2);
//Form3.DBGrid3  - редактирование основной таблицы
if Form1.PopupMenu1.PopupComponent = Form3.DBGrid2 then
  CopyGRDToClb(Form3.DBGrid2);
//DBGrid3  - таилица выборки
if PopupMenu1.PopupComponent = DBGrid3 then
  CopyGRDToClb(DBGrid3);
//DBGrid4 - таблица электроэнергии
if PopupMenu1.PopupComponent = DBGrid4 then
  CopyGRDToClb(DBGrid4);
//DBGrid5  - таблица оплаты коммнулки
if PopupMenu1.PopupComponent = DBGrid5 then
  CopyGRDToClb(DBGrid5);
//DBGrid6 - таблица моего стажа
if PopupMenu1.PopupComponent = DBGrid6 then
  CopyGRDToClb(DBGrid6);
//DBGrid7 - таблица паролей
if PopupMenu1.PopupComponent = DBGrid7 then
  CopyGRDToClb(DBGrid7);

//PopupMenu1.PopupComponent.Name;
//CopyGRDToClb(.PopupComponent.Name);
//CopyGRDToClb(DBGrid6);

end;

procedure TForm1.N11Click(Sender: TObject);
begin
Form1.Close;
end;

procedure TForm1.N12Click(Sender: TObject);
begin
ShellExecute(Application.Handle, 'open', 'C:\windows\system32\osk.exe', nil, nil,SW_SHOWNOACTIVATE);
end;

procedure TForm1.N13Click(Sender: TObject);
begin
Form11.Show;
end;

procedure TForm1.N14Click(Sender: TObject);
begin
ShellExecute(0,'open','Calc.exe','','C:\Windows\system32\',SW_SHOW);
end;

procedure TForm1.N1Click(Sender: TObject);
//var
// Text_sql : string;
// p1      : integer;

begin
{
Text_sql := 'delete from Dengi where id = :p1';
   p1 := DataModule1.FDQuery_Dengi_Zarpl.FieldByName('id').AsInteger;

   DataModule1.FDQuery_Dengi.Active :=false;
   DataModule1.FDQuery_Dengi.SQL.Clear;
   DataModule1.FDQuery_Dengi.SQL.Add(Text_SQL);

   DataModule1.FDQuery_Dengi.Params.ParamByName('p1').Value
    := p1;
   DataModule1.FDQuery_Dengi.ExecSQL;

   Text_Sql := 'select * from Dengi ';
   DataModule1.FDQuery_Dengi.Active :=false;
   DataModule1.FDQuery_Dengi.SQL.Clear;
   DataModule1.FDQuery_Dengi.SQL.Add(Text_SQL);
   DataModule1.FDQuery_Dengi.Active := true;
   DataModule1.FDQuery_Dengi.Last;

 }
end;

procedure TForm1.N2Click(Sender: TObject);
//var
// Text_sql : string;
// p1   :string;
begin
{  Text_Sql := 'delete FROM Dengi  ';
   DataModule1.FDQuery_Dengi.Active :=false;
   DataModule1.FDQuery_Dengi.SQL.Clear;
   DataModule1.FDQuery_Dengi.SQL.Add(Text_SQL);
   DataModule1.FDQuery_Dengi.ExecSQL;

   Text_Sql := 'delete FROM sqlite_sequence WHERE name = :p1';
   p1:=  'Dengi'  ;

   DataModule1.FDQuery_Dengi.Active :=false;
   DataModule1.FDQuery_Dengi.SQL.Clear;
   DataModule1.FDQuery_Dengi.SQL.Add(Text_SQL);

   DataModule1.FDQuery_Dengi.Params.ParamByName('p1').Value
    := p1;
   DataModule1.FDQuery_Dengi.ExecSQL;

   Text_Sql := 'select * from Dengi ';
   DataModule1.FDQuery_Dengi.Active :=false;
   DataModule1.FDQuery_Dengi.SQL.Clear;
   DataModule1.FDQuery_Dengi.SQL.Add(Text_SQL);
   DataModule1.FDQuery_Dengi.Active := true;
   DataModule1.FDQuery_Dengi.Last;
 }

end;

procedure TForm1.N3Click(Sender: TObject);
begin
//
Form1.Memo1.SelectAll;
Form1.Memo1.CopyToClipboard;
ShellExecute(0,'open','Govorilka_cp.exe','-P90 -E "Adult Female #1, American English (TruVoice)" -C',
'',SW_SHOW);

Form1.Memo1.SelStart:=0; // установить начальное положение выделения
Form1.Memo1.SelLength:=0; // установить длину выделенного текста
{
Memo1.ClearSelection; // удалить выделенный текст
Memo1.CutToClipboard; // вырезать в буфер обмена
Memo1.CopyToClipboard; // скопировать в буфер обмена
Memo1.PasteFromClipboard; // вставить из буфера обмена
Memo1.SelText; // эта функция возвращает текстовую строку выделенного текста
}
end;

procedure TForm1.N4Click(Sender: TObject);
//var
// text :string;
begin
//text := Form1.Memo1.SelText; // эта функция возвращает текстовую строку выделенного текста
Clipboard.AsText:= Form1.Memo1.SelText;
ShellExecute(0,'open','Govorilka_cp.exe','-P90 -E "Adult Female #1, American English (TruVoice)" -C',
'',SW_SHOW);

end;

procedure TForm1.N6Click(Sender: TObject);
begin
Form12.Show;
end;

procedure TForm1.N8Click(Sender: TObject);
var
Rlst: LongBool; //результат выполнения
StartUpInfo: TStartUpInfo; //параметры будущего процесса
ProcessInfo: TProcessInformation; //Отслеживание выполнения
CurrDir:string; //текущая папка
ArchiveDir:string; //папку куда складываем архивы
AppToRun:string; //путь до winrar
CommandLine:string; //строка параметров
filename:string; //имя файла
Error:integer; //номер ошибкок
ExitCode: Cardinal; //код завершения
 Sr         : TSearchRec;
 Mask       :string;

 chosenDirectory : string;
 vibor_dir       : string;
 vibor_dir_name  : string;
 Stroka_pr : string;

 path_arch :string;
 Path_1    :string;
 date_file :string;
 my: String;
begin
//Копируем во временный каталог
if SelectDirectory('Выбери папку','D:\Projects_Delphi\New_Bill_Tab',chosenDirectory) then
   begin
     vibor_dir :=chosenDirectory +'\' ;
   end
else
  begin
   ShowMessage('Выбор каталога прервался');
   Exit;
  end;
//Если есть дирректория - удаляем
  if DirectoryExists(Backup_drive+'md') then
  begin
  DeleteDir(Backup_drive+'md');
  end;

 CreateDir(Backup_drive+'md');

//Если есть файл - удаляем
Path_1 := Backup_drive+'md\';
   Mask :='New_Bill_Tab_'+Copy(today_str,1,2)+'_'
   +Copy(today_str,4,2)+'_'
   +Copy(today_str,7,4)+ '*';
if FindFirst(Path_1+Mask, faAnyFile, sr) = 0 then
  begin
  // Должен освободить ресурсы, используемые этими успешными, поисками
    DeleteFile(Path_1+sr.Name);
  end;
  FindClose(sr);
//========================================

//  if FileExists(Backup_drive+'md\\New_Bill_Tab*') then DeleteFile(Backup_drive+'md\New_Bill_Tab*');
//Создаем заново
Path_1 := Backup_drive+'md';
     if System.SysUtils.DirectoryExists(Path_1) then
       begin
         path_arch := Path_1;
       end
     else
       begin
         CreateDir(Path_1);
         path_arch := Path_1;
       end;

CopyDir(vibor_dir, Backup_drive+'md\');

//Определяем текущюю директорию.
//CurrDir:= ExtractFilePath(Application.ExeName);

//Определяем путь к WinRAR.
AppToRun:='C:\Program Files\WinRAR\rar.exe';
//AppToRun:='C:\Windows\System32\notepad.exe';
//AppToRun:=AppToRun+ '/k';
//Проверяем, есть ли в текущей директории программы, папка Archive.
//Если нет - создаем.

//If not DirectoryExists(CurrDir+'Archive') then
//MkDir(CurrDir+'Archive');

//Открываем выбор файла для архивирования.
//Если выбрали, то обрабатываем.
//if OpenDialog1.Execute then
   begin
//В FileName записываем полный путь к выбранному файлу

vibor_dir_name := Copy(vibor_dir,4,length(vibor_dir));
vibor_dir_name:=StringReplace(vibor_dir_name, '\', '', [rfReplaceAll]);

filename:=Backup_drive+'md';
   date_file := '_'+Copy(today_str,1,2)+'_'
   +Copy(today_str,4,2)+'_'
   +Copy(today_str,7,4)   ;

//В CommandLine записываем командную строку
//В нашем случае командная строка должна выглядеть:
//<пробел> a -ep "<путь до архива>" "<путь до файла>"
 comp_name := ReadComputerName;
 backup_name := 'N_B_t'  +'_'+comp_name;
CommandLine:=' "C:\Program Files\WinRAR\rar.exe"  a -ep1 '
 +Backup_drive+'md\'+backup_name+date_file+'.rar '+filename;
// ExecAndWait(CommandLine);

 //CommandLine:= '';
//Заполнение нулями всего StartUpInfo.

FillChar(StartUpInfo, SizeOf(TStartUpInfo), 0);
//После этого выставляем в нем некоторые параметры.

with StartUpInfo do
   begin
    //Содержит количество байтов, занимаемых структурой TStartUpInfo.
    //Обязательно для заполнения. Инициализируйте как SizeOf(TStartUpInfo).
    cb := SizeOf(TStartUpInfo);
    //Содержит набор флагов, позволяющих управлять созданием дочернего процесса.
    //Показываем окно, курсор - часики.
    dwFlags := STARTF_USESHOWWINDOW or STARTF_FORCEONFEEDBACK;
    //Определяет как должно выглядеть окно запущенного приложения.
    //Нормальное отображение
    wShowWindow := SW_SHOWNORMAL;
   end;

   Rlst:= CreateProcess(PChar(AppToRun), PChar(CommandLine),
   nil, nil, false, NORMAL_PRIORITY_CLASS,
   nil, nil, StartUpInfo, ProcessInfo);

//Отслеживаем выполнение.
//Если запуск успешен
   if Rlst then
   with ProcessInfo do
  begin
   //Ждем завершения инициализации.
    WaitForInputIdle(hProcess, INFINITE);
   //Ждем завершения процесса.
   WaitforSingleObject(ProcessInfo.hProcess, INFINITE);
   //Получаем код завершения.
   GetExitCodeProcess(ProcessInfo.hProcess, ExitCode);
    //Закрываем дескриптор процесса.
   CloseHandle(hThread);
   //Закрываем дескриптор потока.
  CloseHandle(hProcess);
  end
//Иначе
else
  begin
    //В случае ошибки - выводим на экран сообщение.
    Error := GetLastError;
    MessageDlg(SysErrorMessage(Error), mtError, [mbOk], 0);
  end;

end;

ShellExecute(Application.Handle, 'open', 'https://drive.google.com/drive/folders/1baYAwkRHE0n29t2I5vsCG-jurEQRCb8Q', nil, nil,SW_SHOWNOACTIVATE);
ShellExecute(Application.Handle, 'open', 'https://cloud.mail.ru/home/New_Bill', nil, nil,SW_SHOWNOACTIVATE);
ShellExecute(Application.Handle, 'open', 'https://disk.yandex.ru/client/disk/New_Bill', nil, nil,SW_SHOWNOACTIVATE);

end;

procedure TForm1.N9Click(Sender: TObject);
var
Rlst: LongBool; //результат выполнения
StartUpInfo: TStartUpInfo; //параметры будущего процесса
ProcessInfo: TProcessInformation; //Отслеживание выполнения
//CurrDir:string; //текущая папка
//ArchiveDir:string; //папку куда складываем архивы
AppToRun :string; //путь до winrar
CommandLine:string; //строка параметров
filename:string; //имя файла
Error:integer; //номер ошибкок
ExitCode: Cardinal; //код завершения

// chosenDirectory : string;
// vibor_dir       : string;
// Stroka_pr : string;
 Sr         : TSearchRec;
 Mask       :string;


 path_arch :string;
 Path_1    :string;
 date_file :string;
 //my: String;
begin
//Если есть дирректория - удаляем
//  if DirectoryExists('D:\md') then DeleteDir('D:\md');
Path_1 := Backup_drive+'md\';
   Mask :='Bill_New_'+Copy(today_str,1,2)+'_'
   +Copy(today_str,4,2)+'_'
   +Copy(today_str,7,4)+ '*';
if FindFirst(Path_1+Mask, faAnyFile, sr) = 0 then
  begin
  // Должен освободить ресурсы, используемые этими успешными, поисками
    DeleteFile(Path_1+sr.Name);
  end;
  FindClose(sr);


//Создаем заново
Path_1 := Backup_drive+'md';
     if System.SysUtils.DirectoryExists(Path_1) then
       begin
         path_arch := Path_1;
       end
     else
       begin
         CreateDir(Path_1);
         path_arch := Path_1;
       end;

  CopyFile( PChar(Database_path), PChar(Backup_drive+'md\Bill_New.db3'),false);

//Определяем путь к WinRAR.
AppToRun :='C:\Program Files\WinRAR\rar.exe';
//AppToRun:='C:\Windows\System32\notepad.exe';
//AppToRun:=AppToRun+ '/k';
//Проверяем, есть ли в текущей директории программы, папка Archive.
//Если нет - создаем.

//If not DirectoryExists(CurrDir+'Archive') then
//MkDir(CurrDir+'Archive');
 filename:=Backup_drive+'md\Bill_New.db3';
   date_file := '_'+Copy(today_str,1,2)+'_'
   +Copy(today_str,4,2)+'_'
   +Copy(today_str,7,4)   ;
//В CommandLine записываем командную строку
//В нашем случае командная строка должна выглядеть:
//<пробел> a -ep "<путь до архива>" "<путь до файла>"
//CommandLine:=' "C:\Program Files\WinRAR\rar.exe"  a -ep1 '+'D:\md\New_Bill_Tab\Bill_New.db3' +date_file+'.rar '+filename;
CommandLine:=' "C:\Program Files\WinRAR\rar.exe"  a '
 +Backup_drive+'md\Bill_New'+ date_file + '.rar '+filename;
//CommandLine:= '';
//Заполнение нулями всего StartUpInfo.
FillChar(StartUpInfo, SizeOf(TStartUpInfo), 0);
//После этого выставляем в нем некоторые параметры.

with StartUpInfo do
begin

//Содержит количество байтов, занимаемых структурой TStartUpInfo.
//Обязательно для заполнения. Инициализируйте как SizeOf(TStartUpInfo).
cb := SizeOf(TStartUpInfo);
//Содержит набор флагов, позволяющих управлять созданием дочернего процесса.
//Показываем окно, курсор - часики.
dwFlags := STARTF_USESHOWWINDOW or STARTF_FORCEONFEEDBACK;
//Определяет как должно выглядеть окно запущенного приложения.
//Нормальное отображение
wShowWindow := SW_SHOWNORMAL;
end;
//Сам запуск WinRAR с параметрами.
//По сути, мы запускаем следующее:
//WinRar.exe a -ep "<путь до архива>" "<путь до файла>"
//showmessage(AppToRun+CommandLine);
//my:= AppToRun+CommandLine;
Rlst:= CreateProcess(PChar(AppToRun), PChar(CommandLine),
nil, nil, false, NORMAL_PRIORITY_CLASS,
nil, nil, StartUpInfo, ProcessInfo);

//Отслеживаем выполнение.
//Если запуск успешен
if Rlst then
with ProcessInfo do
begin
//Ждем завершения инициализации.
WaitForInputIdle(hProcess, INFINITE);
//Ждем завершения процесса.
WaitforSingleObject(ProcessInfo.hProcess, INFINITE);
//Получаем код завершения.
GetExitCodeProcess(ProcessInfo.hProcess, ExitCode);
//Закрываем дескриптор процесса.
CloseHandle(hThread);
//Закрываем дескриптор потока.
CloseHandle(hProcess);
end
//Иначе
else
  begin
//В случае ошибки - выводим на экран сообщение.
Error := GetLastError;
MessageDlg(SysErrorMessage(Error), mtError, [mbOk], 0);
  end;
end;

procedure TForm1.NotePad1Click(Sender: TObject);
begin
ShellExecute(Application.Handle, 'open', '"C:\Program Files\Notepad++\notepad++.exe"', nil, nil,SW_SHOWNOACTIVATE);
end;

procedure TForm1.Notepd1Click(Sender: TObject);
begin
ShellExecute(Application.Handle, 'open', '"C:\Windows\System32\notepad.exe"', nil, nil,SW_SHOWNOACTIVATE);
end;

procedure TForm1.Paint1Click(Sender: TObject);
begin
ShellExecute(Application.Handle, 'open', 'C:\windows\system32\mspaint.exe', nil, nil,SW_SHOWNOACTIVATE);
end;

procedure TForm1.SelectAll1Click(Sender: TObject);

begin
//DBGrid1  - оплата интернета
if PopupMenu1.PopupComponent = DBGrid1 then
  GridSelectAll(DBGrid1);
//DBGrid2  - основная таблица
if PopupMenu1.PopupComponent = DBGrid2 then
  GridSelectAll(DBGrid2);
//DBGrid2  - редактирование основной таблицы
if PopupMenu1.PopupComponent = Form3.DBGrid2 then
  GridSelectAll(Form3.DBGrid2);
//DBGrid3  - таилица выборки
if PopupMenu1.PopupComponent = DBGrid3 then
  GridSelectAll(DBGrid3);
//DBGrid4 - таблица электроэнергии
if PopupMenu1.PopupComponent = DBGrid4 then
  GridSelectAll(DBGrid4);
//DBGrid5  - таблица оплаты коммнулки
if PopupMenu1.PopupComponent = DBGrid5 then
  GridSelectAll(DBGrid5);
//DBGrid6 - таблица моего стажа
if PopupMenu1.PopupComponent = DBGrid6 then
  GridSelectAll(DBGrid6);
//DBGrid7 - таблица паролей
if PopupMenu1.PopupComponent = DBGrid7 then
  GridSelectAll(DBGrid7);


  GridSelectAll(DBGrid1);
end;

procedure TForm1.StringGrid1DrawCell(Sender: TObject; ACol, ARow: Integer;
  Rect: TRect; State: TGridDrawState);
begin
//
{
with StringGrid1.Canvas do
if StringGrid1.Cells[ACol,ARow]='О!' then
  begin
   Font.Color:=clBlack;
   Brush.Color:=clYellow;
   FillRect(Rect);
     WriteText(StringGrid1.Canvas, Rect, 2, 2, StringGrid1.Cells[Col,Row], taCenter);
  end;
  }
end;

procedure TForm1.Button32Click(Sender: TObject);
var
Path       : String;
Sr         : TSearchRec;
F,F_n      : TextFile; //файловая переменная
S,S1,S3,S2,Itog : string;  //текстовая строка
SL              :TStringList;
SLcanel         :TStringList;
SLprogram       :TStringList;
Search,
SearchName: String;
Input,Leng,i : Integer;
id           : Integer;
channel_id,display_name             : String;
start,stop,channel,peredacha        : String;
//что касается проецесса
si : TStartupInfo;
pi : TProcessInformation;
cmdLine : pChar;
app : pChar;
tag1 :string;
tag2 :string;
Text_Sql :string;

fileName   : string;
fileDate   : Integer;
//для вытягивания файла
FileOnNet, LocalFileName: string;
Path_1   :string;
Path_arch :string;
begin

//===========================================================
  FileOnNet := 'http://iptv.netlux.org/epg/xmltv_trk.xml';
 // if DirectoryExists('D:\md') then DeleteDir('D:\md');
//Создаем заново

Path_1 := ExtractFileDir(Application.ExeName)+'\EGP_cur\';
     if System.SysUtils.DirectoryExists(Path_1) then
       begin
         path_arch := Path_1;
       end
     else
       begin
         CreateDir(Path_1);
         path_arch := Path_1;
       end;


  LocalFileName := Path_1+ 'xmltv_trk.xml';

  if GetInetFile(FileOnNet, LocalFileName) = true then
   // ShowMessage('Download xmltv_trk.xml successful')
   begin
     Form1.Edit15.Color := clHotLight;
   end
  else
    ShowMessage('Error in file xmltv_trk.xml download');
//===========================================================

//S := 'http://iptv.netlux.org/epg/xmltv_trk.xml';

Clipboard.AsText := S;
//запускаем внешний процесс и ждём уго завершения
{
 ShellExecute(Application.Handle, 'open', 'C:\Program Files (x86)\WMHelp Software\WMHelp XmlPad\XmlPad.exe',
  'http://iptv.netlux.org/epg/xmltv_trk.xml', nil,SW_SHOWNOACTIVATE);
 }
{
     zeroMemory(@si,sizeof(si));
    si.cb := sizeof(si);

    zeroMemory(@pi,sizeof(pi));

    cmdLine := 0;
    app := 'C:\Program Files (x86)\WMHelp Software\WMHelp XmlPad\XmlPad.exe';

    if not CreateProcess(app,cmdLine,nil,nil,false,0,nil,nil,si,pi ) then begin
        //ShowMessage('fail to create process');
    end
    else begin
        //ShowMessage('process started');
        waitForSingleObject(pi.hProcess,INFINITE);
        //ShowMessage('process closed');

        closehandle(pi.hProcess);
        closeHandle(pi.hThread);
    end;
//Конец обработки внешнегопроцесса
}
SL        := TStringList.Create;
SLcanel   := TStringList.Create;
SLprogram := TStringList.Create;

SL.LoadFromFile(Path_1+ 'xmltv_trk.xml');
//загружаем каналы в SLcanel
Search:='<channel';
Leng:=Length(Search);
for i := 0 to SL.Count - 1 do
  begin
   Input:=Pos(Search, SL.Strings[i]);
    if Input>0 then
    begin
     S:=SL.Strings[i];
     SLcanel.Add(S);
    end;
  end;
//CopyFile('D:\download\xmltv_trk.xml', 'D:\New_Bill_Tab\EPG_NetLux\EGP_Cur\xmltv_trk.xml', false);
//=====================================================
//загружаем программу в SLprogram
Search:='<programme';
Leng:=Length(Search);
for i := 0 to SL.Count - 1 do
  begin
   Input:=Pos(Search, SL.Strings[i]);
    if Input>0 then
    begin
     S:=SL.Strings[i];
     S1:=  Utf8ToAnsi(SL.Strings[i+1]);
     Itog := IntToStr(i)+ S + S1;
     SLprogram.Add(Itog);
    end;
  end;
//======================================================
SL.Free;
//SLcanel.Free;
//SLprogram.Free
//i := SLcanel.Count;
//i := SLprogram.Count;
//============================================



AssignFile(F,Path_1+'\SLprogram.txt'); //связывает F с именем файла
ReWrite(F); //создает новый файл и открывает его
for i := 0 to SLprogram.Count - 1 do
   begin
     Writeln(f,SLprogram.Strings[i]);
   end;
CloseFile(F);

AssignFile(F,Path_1+'\SLcanel.txt'); //связывает F с именем файла
ReWrite(F); //создает новый файл и открывает его
for i := 0 to SLcanel.Count - 1 do
   begin
     Writeln(f,SLcanel.Strings[i]);
   end;
CloseFile(F);
//======================================
//заполняем таблицу EGP_Chanel

Text_Sql := 'DELETE FROM EGP_channel';
DataModule1.FDQuery_C3.Active :=false;
DataModule1.FDQuery_C3.SQL.Clear;
DataModule1.FDQuery_C3.SQL.Add(Text_SQL);
DataModule1.FDQuery_C3.ExecSQL;

Text_Sql := 'DELETE FROM sqlite_sequence WHERE name = :p1';

DataModule1.FDQuery_C3.Active :=false;
DataModule1.FDQuery_C3.SQL.Clear;
DataModule1.FDQuery_C3.SQL.Add(Text_SQL);
DataModule1.FDQuery_C3.Params.ParamByName('p1').Value := 'EGP_channel' ;
DataModule1.FDQuery_C3.ExecSQL;

Text_Sql := 'insert into  EGP_channel (channel_id,display_name) values (:p1,:p2)';
DataModule1.FDQuery_C3.Active :=false;
DataModule1.FDQuery_C3.SQL.Clear;
DataModule1.FDQuery_C3.SQL.Add(Text_SQL);

for i := 0 to SLcanel.Count - 1 do
//channel_id,display_name
     begin
      tag1:='<channel id="';
      tag2:='">';
      channel_id := ParseStr(SLcanel.Strings[i],tag1,tag2);
      tag1:='<display-name lang="ru">';
      tag2:='</display-name>';
      display_name := Utf8ToAnsi(ParseStr(SLcanel.Strings[i],tag1,tag2));

      DataModule1.FDQuery_C3.Params.ParamByName('p1').Value := channel_id ;
      DataModule1.FDQuery_C3.Params.ParamByName('p2').Value := display_name ;
      DataModule1.FDQuery_C3.ExecSQL;
     end;

//==============================================
//start|stop|channel
//Заполнят таблиу EGP_program
Text_Sql := 'DELETE FROM EGP_program';
DataModule1.FDQuery_C3.Active :=false;
DataModule1.FDQuery_C3.SQL.Clear;
DataModule1.FDQuery_C3.SQL.Add(Text_SQL);
DataModule1.FDQuery_C3.ExecSQL;

Text_Sql := 'DELETE FROM sqlite_sequence WHERE name = :p1';

DataModule1.FDQuery_C3.Active :=false;
DataModule1.FDQuery_C3.SQL.Clear;
DataModule1.FDQuery_C3.SQL.Add(Text_SQL);
DataModule1.FDQuery_C3.Params.ParamByName('p1').Value := 'EGP_program' ;
DataModule1.FDQuery_C3.ExecSQL;

Text_Sql := 'insert into  EGP_program (start,stop,channel,peredacha) values (:p1,:p2,:p3,:p4)';
DataModule1.FDQuery_C3.Active :=false;
DataModule1.FDQuery_C3.SQL.Clear;
DataModule1.FDQuery_C3.SQL.Add(Text_SQL);

for i := 0 to SLprogram.Count - 1 do
//channel_id,display_name
     begin

      tag1:='<programme start="';
      tag2:=' +0300" stop="';
      S:= SLprogram.Strings[i];
      start := ParseStr(SLprogram.Strings[i],tag1,tag2);
      start :=Copy(start,1,4)
  +'-'+Copy(start,5,2)
  +'-'+Copy(start,7,2)
  +' '+Copy(start,9,2)
  +':'+Copy(start,11,2)
  +':'+Copy(start,13,2) ;

      tag1:=' +0300" stop="';
      tag2:=' +0300"';
      stop := ParseStr(SLprogram.Strings[i],tag1,tag2);
      stop :=Copy(stop,1,4)
  +'-'+Copy(stop,5,2)
  +'-'+Copy(stop,7,2)
  +' '+Copy(stop,9,2)
  +':'+Copy(stop,11,2)
  +':'+Copy(stop,13,2) ;

      tag1:='channel="';
      tag2:='"';
      channel := ParseStr(SLprogram.Strings[i],tag1,tag2);

      tag1:='<title>';
      tag2:='</t';
      peredacha := ParseStr(SLprogram.Strings[i],tag1,tag2);

      DataModule1.FDQuery_C3.Params.ParamByName('p1').Value := start ;
      DataModule1.FDQuery_C3.Params.ParamByName('p2').Value := stop ;
      DataModule1.FDQuery_C3.Params.ParamByName('p3').Value := channel ;
      DataModule1.FDQuery_C3.Params.ParamByName('p4').Value := peredacha ;
      DataModule1.FDQuery_C3.ExecSQL;

     end;

//ParseStr(str, tag1, tag2: string): tstrings;
//==============================================
SLcanel.Free;
SLprogram.Free;
//==============================================
  fileName := Path_1+'\xmltv_trk.xml';
  fileDate := FileAge(fileName);
   if fileDate > -1 then
   Form1.Edit15.Text :=( 'посл.изм   ' +
     DateToStr(FileDateToDateTime(fileDate)));

     //fileDate := DaysBetween(FileDateToDateTime(fileDate),now);

  if DaysBetween(FileDateToDateTime(fileDate),now) > 1 then
    begin
     Form1.Edit15.Color := clRed ;
    end;

end;


procedure TForm1.Button33Click(Sender: TObject);
var
  dir :string;
begin
SelectDirectory('Путь к рабочей базе', '', dir,[sdShowFiles]);
Edit20.Text := dir;
end;



procedure TForm1.Button34Click(Sender: TObject);
var
Seg             : TDateTime;
data_c          :   string;
Text_sql        :   string;
channel_en      :   string;
channel_ru      :   string;
i               :  integer;
T_start         :   TDateTime;
T_start_s         :   string;
T_start_t         :   string;
T_stop          :   TDateTime;
T_stop_s          :   string;
T_stop_t          :   string;
T_START_STOP      :   string;
peredacha       :   string;
SL              :TStringList;
SL_1            :TStringList;
S               :   string;

begin

Text_sql := 'delete from EPG_CUR';
DataModule1.FDQuery_C3.Active :=false;
DataModule1.FDQuery_C3.SQL.Clear;
DataModule1.FDQuery_C3.SQL.Add(Text_SQL);
DataModule1.FDQuery_C3.ExecSQL;

Text_Sql := 'delete FROM sqlite_sequence WHERE name =:p1';
DataModule1.FDQuery_C3.Active :=false;
DataModule1.FDQuery_C3.SQL.Clear;
DataModule1.FDQuery_C3.SQL.Add(Text_SQL);
DataModule1.FDQuery_C3.Params.ParamByName('p1').Value := 'EPG_CUR' ;
DataModule1.FDQuery_C3.ExecSQL;

Seg := Now;
data_c := FormatDateTime('yyyy-mm-dd hh:nn', Seg);

Text_sql := 'select * from EGP_program  ';
Text_sql :=Text_sql +'WHERE DATETIME(';
Text_sql :=Text_sql + chr(39)+ data_c+chr(39) + ') ';
Text_sql :=Text_sql +'BETWEEN "start" AND "stop"';
Text_sql :=Text_sql + '  order by "channel"';

DataModule1.FDQuery_EPG.Active := false;
DataModule1.FDQuery_EPG.SQL.Clear;
DataModule1.FDQuery_EPG.SQL.Add(Text_sql);
DataModule1.FDQuery_EPG.Active := true;
DataModule1.FDQuery_EPG.Last;
DataModule1.FDQuery_EPG.First;

SL   := TStringList.Create;
SL_1 := TStringList.Create;
for i := 0 to DataModule1.FDQuery_EPG.RecordCount -1 do
  begin
    channel_en :=  DataModule1.FDQuery_EPG.FieldByName('channel').AsString;
    T_start := DataModule1.FDQuery_EPG.FieldByName('start').AsDateTime;
    T_start_s := FormatDateTime('yyyy-mm-dd hh:nn',T_start)  ;
    T_start_t := TimeToStr(T_start);
    T_stop := DataModule1.FDQuery_EPG.FieldByName('stop').AsDateTime ;
    T_stop_s := FormatDateTime('yyyy-mm-dd hh:nn',T_stop)  ;
    T_stop_t := TimeToStr(T_stop);
    T_START_STOP := IntToStr(MinutesBetween(T_start,T_stop));

    peredacha := DataModule1.FDQuery_EPG.FieldByName('peredacha').AsString;
    Text_sql := 'select * from EGP_channel where channel_id = :p0';
     DataModule1.FDQuery_EPG_channel.Active := false;
     DataModule1.FDQuery_EPG_channel.SQL.Clear;
     DataModule1.FDQuery_EPG_channel.SQL.Add(Text_sql);
     DataModule1.FDQuery_EPG_channel.Params.ParamByName('p0').Value := channel_en ;
     DataModule1.FDQuery_EPG_channel.Active := true;
     channel_ru := DataModule1.FDQuery_EPG_channel.FieldByName('display_name').AsString;
       // MinutesBetween(
    S := channel_en + '\\\'+channel_ru+ '\\\'
      +peredacha + '\\\'
      +T_start_s+ '\\\'
      +T_stop_s+ '\\\'
      +T_start_t+ '\\\'
      +T_stop_t+ '\\\'
      +T_START_STOP+ '\\\';
      SL.Add(S);


     DataModule1.FDQuery_EPG.Next;
  end;
DataModule1.FDQuery_EPG.Active := false;
// Massive:= TStringDynArray.Create;
DataModule1.FDQuery_EPG_CUR.Active := false;

for i := 0 to SL.Count - 1 do
  begin
   SL_1 := Tokenize(SL[i], '\\\');
    T_start_stop  := SL_1[0];
    T_stop_t      := SL_1[1];
    T_start_t     := SL_1[2];
    T_stop_s      := SL_1[3];
    T_start_s     := SL_1[4];
    peredacha     := SL_1[5];
    channel_ru    := SL_1[6];
    channel_en    := SL_1[7];

     Text_sql := 'insert into EPG_CUR  (channel_en, channel_ru, peredacha, T_start, T_stop,T_start_time,T_stop_time,T_start_stop)';
     Text_sql :=Text_sql + '  values (:p1,:p2,:p3,:p4,:p5,:p6,:p7,:p8)';
     DataModule1.FDQuery_EPG_CUR.Active := false;
     DataModule1.FDQuery_EPG_CUR.SQL.Clear;
     DataModule1.FDQuery_EPG_CUR.SQL.Add(Text_sql);
     DataModule1.FDQuery_EPG_CUR.Params.ParamByName('p1').Value := channel_en ;
     DataModule1.FDQuery_EPG_CUR.Params.ParamByName('p2').Value := channel_ru ;
     DataModule1.FDQuery_EPG_CUR.Params.ParamByName('p3').Value := peredacha ;
     DataModule1.FDQuery_EPG_CUR.Params.ParamByName('p4').Value := T_start_s ;
     DataModule1.FDQuery_EPG_CUR.Params.ParamByName('p5').Value := T_stop_s ;
     DataModule1.FDQuery_EPG_CUR.Params.ParamByName('p6').Value := T_start_t ;
     DataModule1.FDQuery_EPG_CUR.Params.ParamByName('p7').Value := T_stop_t ;
     DataModule1.FDQuery_EPG_CUR.Params.ParamByName('p8').Value := T_start_stop ;

     DataModule1.FDQuery_EPG_CUR.ExecSQL;

  end;


//Form1.Edit16.Text := '';
//Form1.Edit16.SetFocus;
Text_sql := 'select * from EPG_CUR ';
DataModule1.FDQuery_EPG_CUR.Active := false;
DataModule1.FDQuery_EPG_CUR.SQL.Clear;
DataModule1.FDQuery_EPG_CUR.SQL.Add(Text_sql);
DataModule1.FDQuery_EPG_CUR.Active := true;

Form1.Edit16.Text := 'обновл. '+ TimeToStr(Seg);

DataModule1.FDQuery_EPG_CUR.Active := true;
//Form1.DBGrid9.SetFocus;
Form1.Edit17.Clear;
DataModule1.FDQuery_EPG_CUR.Last;

Form1.Label47.Caption :=  IntToStr(DataModule1.FDQuery_EPG_CUR.RecordCount )
+ '  записей';


//FitGrid(DBGrid9);

end;

procedure TForm1.Button35Click(Sender: TObject);
begin
//%windir%\system32\osk.exe
ShellExecute(Application.Handle, 'open', 'C:\windows\system32\osk.exe', nil, nil,SW_SHOWNOACTIVATE);
end;

procedure TForm1.Button36Click(Sender: TObject);
begin
//
Form1.Memo1.Font.Size := Form1.Memo1.Font.Size +1;
end;

procedure TForm1.Button37Click(Sender: TObject);
begin
//
Form1.Memo1.Font.Size := Form1.Memo1.Font.Size - 1;
end;

procedure TForm1.Button38Click(Sender: TObject);
begin
//
Form1.N3Click(Sender);
end;

procedure TForm1.Button39Click(Sender: TObject);
var
dt1 :TDate;
dt2 :TDate;
Res, dney,a : Integer;
ost_na_seg  : Integer;

MonthsBetween_rasch : integer;
data_tek : TDateTime;
data_tek_str :  string;
god_tek              : integer;
mes_tek              : integer;
DaysInAMonth_tek     : integer;
ejed_day_pay_tek     : real;
neobhod_summa     : real;
//======================
Ostatok_na_konec_tek       : real;
vsego_deneg      : real;
chet_dney : integer;


begin
//Если ввел дату
if Form1.DateTimePicker5.Date <> date() then
 begin
 dt1:= date(); // текущая дата
 dt2:=Form1.DateTimePicker5.Date;
//Возвращает -1 (LessThanValue), если день даты А меньше дня даты В, 0 (EqualsValue), если дни А и В совпадают, +1 (GreaterThanValue), если день даты А больше дня даты В.
 Res:= CompareDate(Dt1, Dt2);
 if Res < 0 then
   begin
    MonthsBetween_rasch := MonthsBetween(dt2,dt1);
   dney := DaysBetween(dt2,dt1);
   neobhod_summa:=0;
    data_tek := dt1;
    god_tek := god;
    mes_tek := mes;


      While data_tek < dt2 do
       begin
        data_tek_str := DateToStr(data_tek);
        god_tek := StrToInt(copy(data_tek_str,7,4));
        mes_tek := StrToInt(copy(data_tek_str,4,2));
        DaysInAMonth_tek := DaysInAMonth(god_tek, mes_tek);
        ejed_day_pay_tek := ejem_platej / DaysInAMonth_tek;
        neobhod_summa := neobhod_summa + ejed_day_pay_tek;
        data_tek := IncDay(data_tek);
       end;

   ost_na_seg := Form1.SpinEdit3.Value;
   Form1.SpinEdit5.Value := Round(neobhod_summa - ost_na_seg) ;


   Form1.Edit21.Text := IntToStr(dney)+'  д ';
   end
 else
   begin
   ShowMessage('До даты меньше 0 дней');
   end;
 end;
if Form1.SpinEdit5.Value <> 0 then
 begin
   //
   vsego_deneg := Form1.SpinEdit3.Value + Form1.SpinEdit5.Value;
   data_tek := date();
   chet_dney := 0;

     data_tek_str := DateToStr(data_tek);
     god_tek := StrToInt(copy(data_tek_str,7,4));
     mes_tek := StrToInt(copy(data_tek_str,4,2));

    While vsego_deneg > ejem_platej do
    begin
     DaysInAMonth_tek := DaysInAMonth(god_tek, mes_tek);
     ejed_day_pay_tek := ejem_platej / DaysInAMonth_tek;
  //   Ostatok_na_konec_tek := vsego_deneg;
     vsego_deneg :=vsego_deneg - (DaysInAMonth_tek*ejed_day_pay_tek) ;

      if vsego_deneg > 0 then
      begin
       Inc(mes_tek);
       chet_dney := chet_dney + DaysInAMonth_tek;
      end;
    end;

   data_tek := IncDay(date(), chet_dney  ); //вычисляем промежуточную дату
//===================================================
//идем по дням
     data_tek_str := DateToStr(data_tek);
     god_tek := StrToInt(copy(data_tek_str,7,4));
     mes_tek := StrToInt(copy(data_tek_str,4,2));

    While vsego_deneg > 0 do
     begin
        DaysInAMonth_tek := DaysInAMonth(god_tek, mes_tek);
        ejed_day_pay_tek := ejem_platej / DaysInAMonth_tek;
        vsego_deneg :=vsego_deneg - ejed_day_pay_tek ;
        Inc(chet_dney);
     end;
   Form1.DateTimePicker5.Enabled := true;
   Form1.DateTimePicker5.Date := IncDay(date(), chet_dney);


 end;
// ejed_day_pay_tek


//
end;

procedure TForm1.Label48Click(Sender: TObject);
begin
//
  if Form1.sortFlag1 then
  //На английский
  begin
    LoadKeyboardLayout(StrCopy(Layout,'00000409'),KLF_ACTIVATE);

    Form1.Label48.Caption:= NameKeyboardLayout(GetActiveKbdLayout);
  end
  else
  begin
    //На русский
    LoadKeyboardLayout( StrCopy(Layout,'00000419'),KLF_ACTIVATE);
    Form1.Label48.Caption:= NameKeyboardLayout(GetActiveKbdLayout);
  end;


Form1.sortFlag1 := not Form1.sortFlag1;
end;

end.
