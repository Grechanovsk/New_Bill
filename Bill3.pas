unit Bill3;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Data.DB, Vcl.Grids, Vcl.DBGrids,
  Vcl.StdCtrls, FireDAC.Stan.Intf, FireDAC.Stan.Option, FireDAC.Stan.Param,
  FireDAC.Stan.Error, FireDAC.DatS, FireDAC.Phys.Intf, FireDAC.DApt.Intf,
  FireDAC.Stan.Async, FireDAC.DApt,
  DateUtils,
  FireDAC.Comp.DataSet, FireDAC.Comp.Client,
  Vcl.ExtCtrls;

type
  TForm3 = class(TForm)
    DBGrid2: TDBGrid;
    Button1: TButton;
    Panel4: TPanel;
    Label6: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    Label15: TLabel;
    Edit1: TEdit;
    Edit2: TEdit;
    Edit3: TEdit;
    Edit5: TEdit;
    Edit4: TEdit;
    Edit6: TEdit;
    ComboBox2: TComboBox;
    ComboBox1: TComboBox;
    Button2: TButton;
    ComboBox3: TComboBox;
    Label1: TLabel;
    CheckBox1: TCheckBox;
    CheckBox2: TCheckBox;
    CheckBox3: TCheckBox;
    CheckBox4: TCheckBox;
    Edit7: TEdit;
    Label2: TLabel;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Button1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure DataSource1DataChange(Sender: TObject; Field: TField);
    procedure DBGrid2DrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure FormActivate(Sender: TObject);
    procedure Button2Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form3: TForm3;

implementation

{$R *.dfm}

uses New_Bill_Tab, DM;


function DateToBase(ish_str: string): string;
begin
Result  := Copy(ish_str,7,4)+ '-'+Copy(ish_str,4,2)+'-'+Copy(ish_str,1,2);
end;

procedure TForm3.Button1Click(Sender: TObject);
begin
  Form3.Close;
end;

procedure TForm3.Button2Click(Sender: TObject);
var
Text_SQL : String;
id       : integer;
day_nedeli    : array[1..7] of string;

//  Trawm           :integer;
//  Zapl_inet       :integer;
//  Gips            :integer;
//  Sn_Gips         :integer;
 // Den_Gizni       :integer;
 // Sobitie     : string;
    DenNed      : string;
 //   Data        : string;
 //   Ostatok     : string;
 //   Pil         : string;

begin
 day_nedeli[1] := 'Понедельник';
 day_nedeli[2] := 'Вторник';
 day_nedeli[3] := 'Среда';
 day_nedeli[4] := 'Четверг';
 day_nedeli[5] := 'Пятница';
 day_nedeli[6] := 'Суббота';
 day_nedeli[7] := 'Воскресенье';

DenNed:= day_nedeli[DayOfTheWeek( StrToDate(Form1.today_str))] ;
id:=DataModule1.FDQuery_Edit.FieldByName('id').AsInteger;
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
Text_SQL :=Text_SQL +'DenNed   = :p10';
//Text_SQL :=Text_SQL +'Data   = :p11';
Text_SQL :=Text_SQL + '  WHERE id = :p12;';

DataModule1.FDQuery_ins.Active :=false;
DataModule1.FDQuery_ins.SQL.Clear;
DataModule1.FDQuery_ins.SQL.Add(Text_SQL);
DataModule1.FDQuery_ins.Params.ParamByName('p1').Value := Form3.Edit7.Text;;
DataModule1.FDQuery_ins.Params.ParamByName('p3').Value := Form3.ComboBox1.Text;
DataModule1.FDQuery_ins.Params.ParamByName('p4').Value := Form3.Edit2.Text;
DataModule1.FDQuery_ins.Params.ParamByName('p5').Value := Form3.Edit3.Text;
DataModule1.FDQuery_ins.Params.ParamByName('p6').Value := Form3.Edit5.Text;
DataModule1.FDQuery_ins.Params.ParamByName('p7').Value := Form3.Edit4.Text;
DataModule1.FDQuery_ins.Params.ParamByName('p8').Value := Form3.ComboBox2.Text;
DataModule1.FDQuery_ins.Params.ParamByName('p9').Value := Form3.Edit6.Text;
DataModule1.FDQuery_ins.Params.ParamByName('p10').Value := Form3.ComboBox3.Text;

DataModule1.FDQuery_ins.Params.ParamByName('p12').Value := IntToStr(id) ;

DataModule1.FDQuery_ins.ExecSQL;

DataModule1.FDQuery_Edit.Active := false;
DataModule1.FDQuery_Edit.Active := true;

   Form3.DBGrid2.Columns[0].Width := 30;
   Form3.DBGrid2.Columns[0].Alignment:=taCenter;
   Form3.DBGrid2.Columns[1].Width := 80;
   Form3.DBGrid2.Columns[2].Width := 80;
   Form3.DBGrid2.Columns[3].Width := 80;
   Form3.DBGrid2.Columns[3].Alignment:=taCenter;
   Form3.DBGrid2.Columns[4].Width := 50;
   Form3.DBGrid2.Columns[4].Alignment:=taCenter;
   Form3.DBGrid2.Columns[5].Width := 80;
   Form3.DBGrid2.Columns[6].Width := 80;
   Form3.DBGrid2.Columns[7].Width := 80;
   Form3.DBGrid2.Columns[8].Width := 80;
   Form3.DBGrid2.Columns[9].Width := 80;
   Form3.DBGrid2.Columns[10].Width := 80;

   DataModule1.DataSource_Edit.DataSet.Locate('id',id,[loCaseInsensitive, loPartialKey]);
//   Form3.DataSource1.DataSet.Locate('Data',Form1.today_str,[loCaseInsensitive, loPartialKey]);

   Form3.DBGrid2.SetFocus;

end;

procedure TForm3.DataSource1DataChange(Sender: TObject; Field: TField);
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

procedure TForm3.DBGrid2DrawColumnCell(Sender: TObject; const Rect: TRect;
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

procedure TForm3.FormActivate(Sender: TObject);
begin
DataModule1.DataSource_Edit.DataSet.Locate('Data',Form1.today_str,[loCaseInsensitive, loPartialKey]);
Form3.DBGrid2.SetFocus;
end;

procedure TForm3.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Form1.Show;
end;

procedure TForm3.FormCreate(Sender: TObject);
var
Text_SQL : String;
//id       : integer;
//day_nedeli    : array[1..7] of string;
begin
DataModule1.FDConnection_osn.DriverName :='SQLite'  ;
DataModule1.FDConnection_osn.Params.Database := ExtractFileDir(Application.ExeName)+'\'+ 'Bill_New.db3';
DataModule1.FDConnection_osn.Connected := true ;

Text_SQL :='select * from Bill_New ';

DataModule1.FDQuery_Edit.Active :=false;
DataModule1.FDQuery_Edit.SQL.Clear;
DataModule1.FDQuery_Edit.SQL.Add(Text_SQL);
DataModule1.FDQuery_Edit.Active :=true;

   Form3.DBGrid2.Columns[0].Width := 30;
   Form3.DBGrid2.Columns[0].Alignment:=taCenter;
   Form3.DBGrid2.Columns[1].Width := 80;
   Form3.DBGrid2.Columns[2].Width := 80;
   Form3.DBGrid2.Columns[3].Width := 80;
   Form3.DBGrid2.Columns[3].Alignment:=taCenter;
   Form3.DBGrid2.Columns[4].Width := 50;
   Form3.DBGrid2.Columns[4].Alignment:=taCenter;
   Form3.DBGrid2.Columns[5].Width := 80;
   Form3.DBGrid2.Columns[6].Width := 80;
   Form3.DBGrid2.Columns[7].Width := 80;
   Form3.DBGrid2.Columns[8].Width := 80;
   Form3.DBGrid2.Columns[9].Width := 80;
   Form3.DBGrid2.Columns[10].Width := 80;

   DataModule1.DataSource_Edit.DataSet.Locate('Data',Form1.today_str,[loCaseInsensitive, loPartialKey]);

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

end.
