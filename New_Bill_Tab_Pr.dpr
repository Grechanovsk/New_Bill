program New_Bill_Tab_Pr;

uses
  Vcl.Forms,
  New_Bill_Tab in 'New_Bill_Tab.pas' {Form1},
  DM in 'DM.pas' {DataModule1: TDataModule},
  Bill2 in 'Bill2.pas' {Form2},
  Bill3 in 'Bill3.pas' {Form3},
  Bill_Opl_Sveta in 'Bill_Opl_Sveta.pas' {Form8},
  Bill_Opl_Sheta in 'Bill_Opl_Sheta.pas' {Form7},
  Bill_Svet in 'Bill_Svet.pas' {Form5},
  Bill_Shet in 'Bill_Shet.pas' {Form6},
  Bill_Pokz_Voda in 'Bill_Pokz_Voda.pas' {Form4},
  Edit_Zam in 'Edit_Zam.pas' {Form9},
  Bill_Login_Sel in 'Bill_Login_Sel.pas' {Form10},
  Calendar in 'Calendar.pas' {Form11},
  Nstroyki in 'Nstroyki.pas' {Form12};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TForm1, Form1);
  Application.CreateForm(TDataModule1, DataModule1);
  Application.CreateForm(TForm2, Form2);
  Application.CreateForm(TForm3, Form3);
  Application.CreateForm(TForm8, Form8);
  Application.CreateForm(TForm7, Form7);
  Application.CreateForm(TForm5, Form5);
  Application.CreateForm(TForm5, Form5);
  Application.CreateForm(TForm6, Form6);
  Application.CreateForm(TForm4, Form4);
  Application.CreateForm(TForm9, Form9);
  Application.CreateForm(TForm10, Form10);
  Application.CreateForm(TForm11, Form11);
  Application.CreateForm(TForm12, Form12);
  Application.Run;
end.
