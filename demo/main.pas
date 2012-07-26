unit main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, A7Rep, Menus;

type
  TMainForm = class(TForm)
    MainMenu: TMainMenu;
    Rep: TA7Rep;
    SimplereportMenu: TMenuItem;
    procedure SimplereportMenuClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  MainForm: TMainForm;

implementation

{$R *.dfm}

procedure TMainForm.SimplereportMenuClick(Sender: TObject);
var
  i : Integer;
  s, sum : Double;
  Kurs_str : string;
  Kurs : Double;
begin
  sum := 0;
  Rep.OpenTemplate(ExtractFilePath(Application.ExeName)+'SimpleReport.xls');
  Rep.PasteBand('Title');
  Rep.SetValue('#VENDOR#','Vendor name');
  Rep.SetValue('#BUY#','Private person');
  Rep.SetValue('#D#','01.03.2012');
  Rep.SetValue('#NOTE#','-');
  Rep.SetComment('#ID#','Здесь мы комментируем если нужно'); // Обязательно делаем комментарий ПЕРЕД тем как вставим значение в ячейку, иначе значение затрет метку и SetComment не найдет куда писать коммент
  Rep.SetValue('#ID#','54321');
  for i :=1 to 10 do begin
    Rep.PasteBand('Line');
    Rep.SetValue('#N#', i);
    Rep.SetValue('#NAME#', 'Item-' + IntToStr(i));
    Rep.SetValue('#UNIT#', 'pkg');
    Rep.SetValue('#QUANT#', i*1.9);
    s := (20-i)*19;
    sum := sum + s;
    Rep.SetValue('#SUMMA#', s);
  end;
  Rep.PasteBand('Foot');
  Rep.SetValue('#SUMMA#', sum);
  Rep.SetValue('#CURRENT_DATE#', DateToStr(now));

  // Пример использование комментария для чтения из него вспомогательной информации
  // К примеру здесь если курс должен быть отличен от дефолтного, то можно его прописать в комментарий к ячейке с меткой #SUMMA_VAL#
  Kurs_str := Rep.GetAndClearComment('#SUMMA_VAL#');
  if Kurs_str<>'' then
    Kurs := StrToFloat(Kurs_str)
  else
    Kurs := 8;
  Rep.SetValue('#SUMMA_VAL#', sum/Kurs);

  Rep.Show;
end;

end.
