//****************************************************************************//
//****** Author - Kucher Alexander Vasilyevich <a7exander@gmail.com> *********//
//************************* (�) 2015 *****************************************//
//****************************************************************************//
unit A7Rep;

interface

uses
  SysUtils, Classes, Variants, ComObj, StdCtrls, Forms;

const
  // --------------- Excel constants ----------------------------
  xlFormulas = $FFFFEFE5;
  xlComments = $FFFFEFD0;
  xlValues = $FFFFEFBD;

type
  // progress-form
  TA7Progress = class
  private
    F: TForm;
    L1: TLabel;
    L2: TLabel;
  protected
  public
    procedure Line(p: integer);
    constructor Create(AOwner: TComponent);
    destructor Destroy; override;
  end;

  TA7Rep = class(TComponent)
  private
    Progress: TA7Progress;

    CurrentLine: integer;
  protected
  public
    Excel, TemplateSheet: Variant;
    FirstBandLine, LastBandLine: integer; // last paste band position
    MaxBandColumns : integer ; // Max band width in excel-columns
    isVisible : boolean;
    procedure OpenTemplate(FileName: string); overload;
    procedure OpenTemplate(FileName: string; Visible : boolean); overload;
    procedure OpenWorkSheet(Name: string);
    procedure PasteBand(BandName: string);
    procedure SetValue(VarName: string; Value: Variant); overload;
    procedure SetValue(X,Y: Integer; Value: Variant); overload;
    procedure SetValueAsText(varName: string; Value: string); overload;
    procedure SetValueAsText(X,Y: Integer; Value: string); overload;    
    procedure SetComment(VarName: string; Value: Variant);
    function GetComment(VarName: string): string;
    function GetAndClearComment(VarName: string): string; overload;
    function GetAndClearComment(X,Y: Integer): string; overload;
    function GetCellName(aCol, aRow : Integer): string;
    procedure ExcelFind(const Value: string; var aCol, aRow : Integer; Where:Integer);
    procedure Show;
    destructor Destroy; override;
  published
  end;

procedure Register;

implementation

const
  MaxBandLines = 300; // max template lines count

procedure Register;
begin
  RegisterComponents('A7rexcel', [TA7Rep]);
end;

{ TWcReport }

destructor TA7Rep.Destroy;
begin

  inherited;
end;

function TA7Rep.GetCellName(aCol, aRow : Integer): string;
var
  c1, c2 : Integer;
begin
  if aCol<27 then begin
    Result := chr(64 + aCol)+IntToStr(aRow);
  end else begin
    c1 := (aCol-1) div 26;
    c2 := aCol - c1*26;
    Result := chr(64 + c1)+chr(64 + c2)+IntToStr(aRow);
  end;
end;

procedure TA7Rep.ExcelFind(const Value: string; var aCol, aRow : Integer; Where:Integer);
// Where: the search type (xlFormulas, xlComments, xlValues)
var
  R: Variant;
begin
   R := TemplateSheet.Rows[IntToStr(FirstBandLine) + ':' + IntToStr(LastBandLine)];
   R:=R.Find(What:=Value,LookIn:=Where);
   if VarIsClear(R) then begin
     aCol := -1;
     aRow := -1;
   end else begin
     aCol:=R.Column;
     aRow:=R.Row;
   end;
end;

function TA7Rep.GetAndClearComment(VarName: string): string;
var
  v : Variant;
  x , y : Integer;
begin
  Result := '';
  ExcelFind(VarName, x, y, xlValues);
  if x<0 then Exit;
  v := Variant(TemplateSheet.Cells[y, x].Value);
  if ((varType(v) = varOleStr)) then
     if Pos(VarName,v)>0 then begin
        Result := TemplateSheet.Cells[y, x].Comment.Text;
        TemplateSheet.Cells[y, x].ClearComments;
     end;
end;

function TA7Rep.GetAndClearComment(X, Y: Integer): string;
var
  v : Variant;
begin
  Result := '';
  if x<0 then Exit;
  v := Variant(TemplateSheet.Cells[y, x].Value);
  if ((varType(v) = varOleStr)) then begin
    Result := TemplateSheet.Cells[y, x].Comment.Text;
    TemplateSheet.Cells[y, x].ClearComments;
  end;
end;

function TA7Rep.GetComment(VarName: string): string;
var
  v : Variant;
  x , y : Integer;
begin
  Result := '';
  ExcelFind(VarName, x, y, xlValues);
  if x<0 then Exit;
  v := Variant(TemplateSheet.Cells[y, x].Value);
  if ((varType(v) = varOleStr)) then
     if Pos(VarName,v)>0 then begin
        Result := TemplateSheet.Cells[y, x].Comment.Text;
     end;
end;

procedure TA7Rep.OpenTemplate(FileName: string);
begin // �� ��������� �� ���������� Excel
  OpenTemplate(FileName, false);
end;

procedure TA7Rep.OpenTemplate(FileName: string; Visible: boolean);
begin
  Excel := CreateOleObject('Excel.Application');
  Excel.Workbooks.Open(FileName, True, True);
  Excel.DisplayAlerts := False; // for prevent error in SetValue procedure, where VarName not fount for replace

  Excel.Visible := Visible;
  isVisible := Visible;
  if isVisible=False then
    Progress := TA7Progress.Create(Self); // show progress-window
  Application.ProcessMessages;

  TemplateSheet := Excel.Workbooks[1].Sheets[1];
  CurrentLine := 1;
  MaxBandColumns := TemplateSheet.UsedRange.Columns.Count;
end;

procedure TA7Rep.PasteBand(BandName: string);
var
  i: integer;
  v, Range: Variant;
begin

  // find band in template
  FirstBandLine := 0; LastBandLine := 0;
  i := CurrentLine;
  while ((LastBandLine = 0) and (i < CurrentLine + MaxBandLines)) do begin
    v := Variant(TemplateSheet.Cells[i, 1].Value);
    if (varType(v) = varOleStr) and (FirstBandLine = 0) then begin
      if v = BandName then begin // the start of band
        FirstBandLine := i;
      end;
    end;
    if (FirstBandLine <> 0) then begin
      if not ((varType(v) = varOleStr) and (v = BandName)) then LastBandLine := i - 1;
    end;
    inc(i);
  end;

  if LastBandLine>0 then begin // if BandName found
    Range := TemplateSheet.Rows[IntToStr(FirstBandLine) + ':' + IntToStr(LastBandLine)];
    Range.Copy;
    Range := TemplateSheet.Rows[IntToStr(CurrentLine) + ':' + IntToStr(CurrentLine)];
    Range.Insert;

    // delete band name from result lines
    for i := CurrentLine to CurrentLine + (LastBandLine - FirstBandLine) do begin
      TemplateSheet.Cells[i, 1].Value := '';
    end;
    CurrentLine := CurrentLine + (LastBandLine - FirstBandLine) + 1;
    // new band position in report
    FirstBandLine := CurrentLine - (LastBandLine - FirstBandLine) - 1;
    LastBandLine := CurrentLine - 1;
    if isVisible=false then
      Progress.Line(CurrentLine);
  end;
end;

procedure TA7Rep.SetComment(VarName: string; Value: Variant);
// VarName - tag in cell, for setting comment
var
  x, y : Integer;
begin
  ExcelFind(VarName, x, y, xlValues);
  if x>0 then begin
    TemplateSheet.Cells[y, x].AddComment(Value);
  end;
end;

procedure TA7Rep.SetValue(VarName: string; Value: Variant);
var Range: Variant;
  s: string;
begin
  if Value=null then
    s := ''
  else
    s := Value;
  Range := TemplateSheet.Rows[IntToStr(FirstBandLine) + ':' + IntToStr(LastBandLine)];
  Range.Replace(VarName, s);
end;

procedure TA7Rep.SetValue(X, Y: Integer; Value: Variant);
var
  s: string;
begin
  if Value=null then
    s := ''
  else
    s := Value;
  TemplateSheet.Cells[y, x].Value := s;
end;

procedure TA7Rep.SetValueAsText(varName, Value: string);
var y, x: integer;
  v: Variant;
begin
  for y := FirstBandLine to LastBandLine do begin
    for x := 2 to MaxBandColumns do begin
      v := Variant(TemplateSheet.Cells[y, x].Value);
      if ((varType(v) = varOleStr)) then
        if v = VarName then begin
          TemplateSheet.Cells[y, x].NumberFormat:= '@';
          TemplateSheet.Cells[y, x].Value := Value;
        end;
    end;
  end;
end;

procedure TA7Rep.SetValueAsText(X, Y: Integer; Value: string);
begin
  TemplateSheet.Cells[y, x].NumberFormat:= '@';
  TemplateSheet.Cells[y, x].Value := Value;
end;

procedure TA7Rep.Show;
var
  Range: Variant;
begin
  // delete the template from result report
  Range := TemplateSheet.Rows[IntToStr(CurrentLine) + ':' + IntToStr(CurrentLine + MaxBandLines)];
  Range.Delete;

  Excel.Visible := true;
  if isVisible=false then
    Progress.Free;

  Range := Unassigned;
  TemplateSheet := Unassigned;
  Excel := Unassigned;
end;

{ TA7ProgressForm }

constructor TA7Progress.Create(AOwner: TComponent);
begin
  F := TForm.Create(nil);
  F.Width := 150;
  F.Height := 80;
  F.Position := poScreenCenter;
  F.BorderStyle := bsNone;
  F.FormStyle := fsStayOnTop;
  F.Color := $FFFFFF;
  L1 := TLabel.Create(F);
  L1.Parent := F;
  L1.Left := 15;
  L1.Width := 100;
  L1.Alignment := taCenter;
  L1.Top := 20;
  L1.Caption := 'Building report';
  L2 := TLabel.Create(F);
  L2.Parent := F;
  L2.Left := 30;
  L2.Width := 100;
  L2.Alignment := taCenter;
  L2.Top := 40;
  F.Show;
end;

destructor TA7Progress.Destroy;
begin
  L1.Free;
  L2.Free;
  F.Free;
end;

procedure TA7Progress.Line(p: integer);
begin
  L2.Caption := 'Line: ' + IntToStr(p);
  Application.ProcessMessages;
end;

// for using multi-sheet reports
// if Name='' then open first worksheet
procedure TA7Rep.OpenWorkSheet(Name: string);
var
  Range: Variant;
begin
  // delete band templates
  Range := TemplateSheet.Rows[IntToStr(CurrentLine) + ':' + IntToStr(CurrentLine + MaxBandLines)];
  Range.Delete;

  if Name='' then
    TemplateSheet := Excel.Workbooks[1].Sheets[1]
  else
    TemplateSheet := Excel.Workbooks[1].Sheets[Name];
  CurrentLine := 1;
  MaxBandColumns := TemplateSheet.UsedRange.Columns.Count;
end;

end.

