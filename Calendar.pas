unit Calendar;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms,DateUtils,
  // чтобы не было хинтов
  FireDAC.Stan.Param,
  //
  Vcl.Dialogs, Vcl.StdCtrls, Vcl.ComCtrls,
  Vcl.Menus;

type
  TForm11 = class(TForm)
//  {$METHODINFO ON}
    Button1: TButton;
    MonthCalendar1: TMonthCalendar;
    Label29: TLabel;
    Edit9: TEdit;
    Edit1: TEdit;
    Label1: TLabel;
    PopupMenu1: TPopupMenu;
    N1: TMenuItem;
    Label2: TLabel;
    Edit2: TEdit;
    N2: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    Button2: TButton;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Button1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure MonthCalendar1Click(Sender: TObject);
    procedure N1Click(Sender: TObject);
    procedure N2Click(Sender: TObject);
    procedure MonthCalendar1DblClick(Sender: TObject);
    procedure PopupMenu1Popup(Sender: TObject);
    procedure N3Click(Sender: TObject);
    procedure N4Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
//  {$METHODINFO OFF}
  private
    { Private declarations }
  public
  SrPopUP  : boolean;
  SelDate     : TDate;
  SelDate_s   : string;
  zamet    : string;
  s,po        : TDate;
  Text_sql : string;
    { Public declarations }
  end;

var
  Form11: TForm11;

type
   TExecute = procedure of object;

implementation

{$R *.dfm}

uses New_Bill_Tab,DM ,Edit_Zam;

procedure TForm11.Button1Click(Sender: TObject);
begin
Form11.Close;
end;

procedure TForm11.Button2Click(Sender: TObject);
begin
//
s := StrToDate(Form11.Edit1.Text);
po := StrToDate(Form11.Edit2.Text);
if s > po then
   begin
    Form11.Edit9.Text :=' - '+IntToStr(DaysBetween(s, po));
//    Form11.Label29.Caption := 'До выбранного дня осталось:';
   end
 else
   begin
    Form11.Edit9.Text :=IntToStr(DaysBetween(s, po));
//    Form11.Label29.Caption := 'С выбранного дня прошло:';
   end;

end;

procedure TForm11.FormClose(Sender: TObject; var Action: TCloseAction);
begin
Form1.Show;
end;

procedure TForm11.FormCreate(Sender: TObject);
var
StartDate  : TDate;

begin
StartDate := StartOfAYear(CurrentYear);
Form11.MonthCalendar1.Date := StartDate ;
SrPopUP := false;
Form11.Edit1.Text := DateToStr(Now);
Form11.Label1.Caption := 'Сегодня :';

end;

procedure TForm11.MonthCalendar1Click(Sender: TObject);
//var
//SelDate     : TDate;

begin
{//
if SrPopUP = false then
 begin
 SelDate := Form11.MonthCalendar1.Date;
 s := SelDate;
 Form11.Label1.Caption := 'C:';
 Form11.Edit1.Text := DateToStr(Seldate);
 SrPopUP := true;
 n1.Enabled := false;
 end
else
 begin
 SelDate := Form11.MonthCalendar1.Date;
 Form11.Label2.Caption := 'По :';
 Form11.Edit2.Text := DateToStr(Seldate);
 SrPopUP := false;
 n1.Enabled := true;
 Form11.Label29.Caption := 'Дня(Дней)';
  po := SelDate;
  if s > po then
     begin
      Form11.Edit9.Text :=' - '+IntToStr(DaysBetween(s, po));
//    Form11.Label29.Caption := 'До выбранного дня осталось:';
     end
   else
     begin
      Form11.Edit9.Text :=IntToStr(DaysBetween(s, po));
//      Form11.Label29.Caption := 'С выбранного дня прошло:';
     end;

 end;

}
end;

procedure TForm11.MonthCalendar1DblClick(Sender: TObject);
begin
//

if SrPopUP = false then
 begin
 SelDate := Form11.MonthCalendar1.Date;
 s := SelDate;
 Form11.Label1.Caption := 'C:';
 Form11.Edit1.Text := DateToStr(Seldate);
 SrPopUP := true;
 n1.Enabled := false;
 end
else
 begin
 SelDate := Form11.MonthCalendar1.Date;
 Form11.Label2.Caption := 'По :';
 Form11.Edit2.Text := DateToStr(Seldate);
 SrPopUP := false;
 n1.Enabled := true;
 Form11.Label29.Caption := 'Дня(Дней)';
po := SelDate;
if s > po then
   begin
    Form11.Edit9.Text :=' - '+IntToStr(DaysBetween(s, po));
//    Form11.Label29.Caption := 'До выбранного дня осталось:';
   end
 else
   begin
    Form11.Edit9.Text :=IntToStr(DaysBetween(s, po));
//    Form11.Label29.Caption := 'С выбранного дня прошло:';
   end;

 end;



//ExecuteRoutine(Form11, 'PopupMenu1Popup');
// Form11.PopupMenu1Popup(Sender);

end;

procedure TForm11.N1Click(Sender: TObject);
begin
SelDate_s :=  FormatDateTime('yyyy-mm-dd',Form11.MonthCalendar1.Date);

Text_sql := 'select * from Bill_New where Data = :p1';
Form9.FDQuery1.Active := false;
Form9.FDQuery1.SQL.Clear;
Form9.FDQuery1.SQL.Add(Text_SQL);
Form9.FDQuery1.Params.ParamByName('p1').Value := SelDate_s ;
Form9.FDQuery1.Active :=true;

Form9.zamet := Form9.FDQuery1.FieldByName('Zam').AsString;
Form9.kkk  :=  Form9.FDQuery1.FieldByName('id').AsInteger;
Form9.Memo1.Lines.Clear;
Form9.Memo1.Text :=  Form9.zamet ;
Form9.show;
end;

procedure TForm11.N2Click(Sender: TObject);
begin
SelDate_s :=  FormatDateTime('yyyy-mm-dd',Form11.MonthCalendar1.Date);

Text_sql := 'select * from Bill_New where Data = :p1';
Form9.FDQuery1.Active := false;
Form9.FDQuery1.SQL.Clear;
Form9.FDQuery1.SQL.Add(Text_SQL);
Form9.FDQuery1.Params.ParamByName('p1').Value := SelDate_s ;
Form9.FDQuery1.Active :=true;


Form9.zamet := Form9.FDQuery1.FieldByName('Zam').AsString;
Form9.kkk  :=  Form9.FDQuery1.FieldByName('id').AsInteger;
Form9.Memo1.Lines.Clear;
Form9.Memo1.Text :=  Form9.zamet ;
Form9.show;

end;

procedure TForm11.N3Click(Sender: TObject);
begin
 SelDate := Form11.MonthCalendar1.Date;
 s := SelDate;
 Form11.Label1.Caption := 'C:';
 Form11.Edit1.Text := DateToStr(Seldate);
end;

procedure TForm11.N4Click(Sender: TObject);
begin
 SelDate := Form11.MonthCalendar1.Date;
 s := SelDate;
 Form11.Edit2.Text := DateToStr(Seldate);
end;

procedure TForm11.PopupMenu1Popup(Sender: TObject);
begin
//
end;

end.
