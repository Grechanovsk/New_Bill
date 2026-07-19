unit DelphiNotesAge;
 
interface
 
type
  // склонение числа - 0 (ноль), 1 (один) или 2 (два)
  TNumDeclination = (ndZero, ndOne, ndTwo);
 
// возвращает склонение числа
function GetNumDeclination(Num: Integer): TNumDeclination;
 
type
  T3Strs = array [TNumDeclination] of string;
 
const
  sYears: T3Strs = ('лет', 'год', 'года');
  sMonths: T3Strs = ('мес€цев', 'мес€ц', 'мес€ца');
  sDays: T3Strs = ('дней', 'день', 'дн€');
 
  sShortYears: T3Strs = ('л.', 'г.', 'г.');
  sShortMonths: T3Strs = ('м.', 'м.', 'м.');
  sShortDays: T3Strs = ('д.', 'д.', 'д.');
 
type
  TAgeFormat = (afFull, afRoundDays, afRoundMonths);
  // afFull         : 15 лет 7 мес€цев 24 дн€
  // afRoundDays    : 15 лет 8 мес€цев
  // afRoundMonths  : 16 лет
 
// считает возраст (количество лет, мес€цев и дней между Birthday и Today)
procedure DecodeAge(const Birthday, Today: TDateTime;
                var Years, Months, Days: Word);
// форматирует возраст в виде строки
function FormatAge(const Birthday, Today: TDateTime; const sY, sM, sD: T3Strs;
                Format: TAgeFormat = afFull): string; overload;
function FormatAge(const Birthday, Today: TDateTime; Short: Boolean = False;
                Format: TAgeFormat = afFull): string; overload;
 
implementation
 
uses
  SysUtils;
 
function GetNumDeclination(Num: Integer): TNumDeclination;
begin
  num := num mod 100;
  if num > 20 then
    num := num mod 10;
 
  case num of
    1:        Result := ndOne;
    2, 3, 4:  Result := ndTwo;
    else      Result := ndZero;  //0, 5..20:
  end;
end;
 
procedure DecodeAge(const Birthday, Today: TDateTime; var Years, Months, Days: Word);
var
  BY, BM, BD,
  TY, TM, TD: Word;
  dY, dM, dD: Integer;
begin
  if Birthday <= Today then
  begin
    DecodeDate(Birthday, BY, BM, BD);
    DecodeDate(Today, TY, TM, TD);
  end else
  begin
    DecodeDate(Today, BY, BM, BD);
    DecodeDate(Birthday, TY, TM, TD);
  end;
 
  dY := TY - BY;
  dM := TM - BM;
  dD := TD - BD;
 
  if dD < 0 then
  begin
    // отнимаем один мес€ц
    dec(dM);
 
    // корректируем кол-во дней: добавл€ем кол-во дней предыдущего мес€ца (чтобы "выйти из нул€")
    dec(TM);
    if TM = 0 then
      TM := 12;
    dD := dD + MonthDays[IsLeapYear(TY), TM];
//вот здесь ошибка   1.03.2023 dD = -1
//поэтому корректируем  !!!!!!!!!!!!
    if dD <0 then dD :=0;

  end;
  if dM < 0 then
  begin
    // отнимаем 1 год
    dec(dY);
    // корректируем кол-во мес€цев: добавл€ем 12 мес€цев (что составл€ет один год)
    dM := dM + 12;
  end;
  // if dY < 0 then raise ERangeError.Create('');
 
  Years := dY;
  Months := dM;
  Days := dD;
end;
 
function FormatAge(const Birthday, Today: TDateTime; const sY, sM, sD: T3Strs;
                Format: TAgeFormat = afFull): string;
  function Add(const Left, Right: string): string;
  begin
    if Left <> ''
      then Result := Left + ' ' + Right
      else Result := Right;
  end;
var
  Y, M, D: Word;
begin
  DecodeAge(Birthday, Today, Y, M, D);
 
  if Format in [afRoundDays, afRoundMonths] then
  begin
    // Round Date To Month:
    if D >= 15 then
    begin
      inc(M);
      if M >= 12 then
      begin
        M := M - 12;
        inc(Y);
      end;
    end;
    D := 0;
 
    if Format = afRoundMonths then
    begin
      // Round Date To Year:
      if M >= 6 then
      begin
        inc(Y);
      end;
      M := 0;
    end;
  end;
 
  Result := '';
  if Y > 0 then
    Result := Add(IntToStr(Y), sY[GetNumDeclination(Y)]);
  if M > 0 then
    Result := Add(Result, Add(IntToStr(M), sM[GetNumDeclination(M)]));
  if D > 0 then
    Result := Add(Result, Add(IntToStr(D), sD[GetNumDeclination(D)]));
end;
 
function FormatAge(const Birthday, Today: TDateTime; Short: Boolean = False;
                Format: TAgeFormat = afFull): string;
begin
  if Short
    then Result := FormatAge(Birthday, Today, sShortYears, sShortMonths, sShortDays, Format)
    else Result := FormatAge(Birthday, Today, sYears, sMonths, sDays, Format);
end;
 
end.