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
begin
  sum := 0;
  Rep.OpenTemplate(ExtractFilePath(Application.ExeName)+'SimpleReport.xls');
  Rep.PasteBand('Title');
  Rep.SetValue('#VENDOR#','Vendor name');
  Rep.SetValue('#BUY#','Private person');
  Rep.SetValue('#D#','01.03.2012');
  Rep.SetValue('#NOTE#','-');
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
  Rep.Show;
end;

end.
