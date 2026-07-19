unit Nstroyki;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs,
  FileCtrl,
  //for inifiles
  inifiles,
  Vcl.StdCtrls, Vcl.ComCtrls, Vcl.Samples.Spin;

type
  TForm12 = class(TForm)
    Label51: TLabel;
    Edit20: TEdit;
    Button33: TButton;
    ComboBoxEx4: TComboBoxEx;
    Label50: TLabel;
    SpinEdit4: TSpinEdit;
    Label52: TLabel;
    Button15: TButton;
    procedure Button33Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure Button15Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form12: TForm12;

implementation

{$R *.dfm}

uses New_Bill_Tab;

procedure TForm12.Button15Click(Sender: TObject);
var
  Ini: Tinifile; //необходимо создать объект, чтоб потом с ним работатьbegin
begin
   Ini:=TiniFile.Create(extractfilepath(paramstr(0))+'My_Ini.ini');
     if Form12.ComboBoxEx4.Text = '' then
  begin
  // Form1.ComboBoxEx4.Text := 'D:\';
  ShowMessage('Ты не установил диск для архивов перейди на вкладку Set, и установи все нужные значения.');
  //Form1.PageControl1.ActivePage := TabSheet10;
  exit;
  end
else
  begin
  Ini.WriteString('Main','Drive_Backup',Form1.Backup_drive);
  end;
//---------------------------------------------------
  if Form12.Edit20.Text = '' then
  begin
  // Form1.ComboBoxEx4.Text := 'D:\';
  ShowMessage('Ты не установил путь для рабочей DB перейди на вкладку Set, и установи все нужные значения.');
//  Form1.PageControl1.ActivePage := TabSheet10;
  exit;
  end
else
  begin
  Ini.WriteString('Main','Database_path',
    Form12.Edit20.Text);
  end;
//-------------------------------------------

  if Form12.SpinEdit4.Value = 0 then
  begin
  // Form1.ComboBoxEx4.Text := 'D:\';
  ShowMessage('Ты не установил ежемесячную оплату.');
//  Form1.PageControl1.ActivePage := TabSheet10;
  exit;
  end
else
  begin
  Ini.WriteInteger('Main','Ежемесячный платёж',Form1.SpinEdit4.Value);
  end;
Ini.Free;
end;

procedure TForm12.Button33Click(Sender: TObject);
var
  dir :string;
begin
   SelectDirectory('Выбор рабочей базы', '', dir,[sdShowFiles]);
Edit20.Text := dir;

end;

procedure TForm12.FormShow(Sender: TObject);
var
  Ini: Tinifile;
  w: dword;
  Root: string;
  i,DiskType: integer;
  f:textfile; // Dalee textovie peremennie
  Present: TDateTime;      // текущая дата и время
  Year, Month, Day : Word; // год, месяц и число, как
 begin
    w := GetLogicalDrives;
  Root := '#:\';
  for i := 0 to 25 do
  begin
    Root[1] := Char(Ord('A') + i);
    if (W and (1 shl i)) > 0 then
      if (GetDriveType(Pchar(Root)) = DRIVE_FIXED)  then
    begin
      Form12.ComboBoxEx4.Items.Add(Root);
    end;
  Ini:=TiniFile.Create(extractfilepath(paramstr(0))+'My_Ini.ini');
  Form1.Backup_drive := Ini.ReadString('Main','Drive_Backup',Form1.ComboBoxEx1.Text);
  Form1.Database_path := Ini.ReadString('Main','Database_path',Form1.ComboBoxEx1.Text);
  Form1.ejem_platej   := Ini.ReadInteger('Main','Ежемесячный платёж',Form1.SpinEdit4.Value);


  Form12.ComboBoxEx4.Text := Form1.Backup_drive;
  Form12.Edit20.Text := Form1.Database_path;
  Form12.SpinEdit4.Value := Form1.ejem_platej;
  end;
 end;

end.
