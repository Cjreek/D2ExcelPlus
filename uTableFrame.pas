unit uTableFrame;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.Grids, Vcl.ComCtrls,
  Vcl.ExtCtrls, Vcl.StdCtrls, Vcl.Samples.Spin, Vcl.Menus, Clipbrd;

type
  TTableFrame = class(TFrame)
    {$region 'Components'}
    sgTable: TStringGrid;
    Panel1: TPanel;
    cbFixColumns: TCheckBox;
    seFixedColumns: TSpinEdit;
    pmGrid: TPopupMenu;
    miGridAddNew: TMenuItem;
    miGridInsertNew: TMenuItem;
    miGridDeleteRow: TMenuItem;
    miGridAddCopy: TMenuItem;
    miGridInsertCopy: TMenuItem;
    N1: TMenuItem;
    N2: TMenuItem;
    {$endregion}
    {$region 'Eventhandler'}
    procedure sgTableMouseWheelDown(Sender: TObject; Shift: TShiftState;
      MousePos: TPoint; var Handled: Boolean);
    procedure sgTableMouseWheelUp(Sender: TObject; Shift: TShiftState;
      MousePos: TPoint; var Handled: Boolean);
    procedure sgTableKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure sgTableGetEditText(Sender: TObject; ACol, ARow: Integer;
      var Value: string);
    procedure cbFixColumnsClick(Sender: TObject);
    procedure miGridAddNewClick(Sender: TObject);
    procedure miGridAddCopyClick(Sender: TObject);
    procedure miGridDeleteRowClick(Sender: TObject);
    procedure sgTableContextPopup(Sender: TObject; MousePos: TPoint;
      var Handled: Boolean);
    procedure miGridInsertNewClick(Sender: TObject);
    procedure miGridInsertCopyClick(Sender: TObject);
    procedure sgTableSelectCell(Sender: TObject; ACol, ARow: Integer;
      var CanSelect: Boolean);
    procedure sgTableDrawCell(Sender: TObject; ACol, ARow: Integer; Rect: TRect;
      State: TGridDrawState);
    {$endregion}
  private
    FFilename: String;
    FOldValue: String;
    FModified: Boolean;
    FOnModifiedChanged: TNotifyEvent;
    procedure DeleteCurrentRow;
    procedure SetModified(AModified: Boolean);
    function GetTabsheet: TTabsheet;
  public
    constructor Create(AOwner: TComponent; AFilename: String); reintroduce;
    procedure LoadFile(const AFilename: String);
    procedure SaveFile();

    function Find(AText: String; AStartPos: TPoint; AOptions: TFindOptions; out ResultPosition: TPoint): Boolean;
    procedure Select(ACol, ARow: Integer; AScrollTo: Boolean = false);

    property Filename: String read FFilename;
    property Modified: Boolean read FModified;
    property Tabsheet: TTabsheet read GetTabsheet;

    property OnModifiedChanged: TNotifyEvent read FOnModifiedChanged write FOnModifiedChanged;
  end;

implementation

{$R *.dfm}

uses
  Math, StrUtils, System.Types, System.IOUtils;

{ TTableFrame }

procedure TTableFrame.cbFixColumnsClick(Sender: TObject);
begin
  seFixedColumns.Enabled := cbFixColumns.Checked;
  if cbFixColumns.Checked then
    sgTable.FixedCols := seFixedColumns.Value + 1
  else
    sgTable.FixedCols := 1;
end;

constructor TTableFrame.Create(AOwner: TComponent; AFilename: String);
begin
  inherited Create(AOwner);
  Name := TPath.GetFileNameWithoutExtension(AFilename);
  LoadFile(AFilename);
end;

procedure TTableFrame.DeleteCurrentRow;
var currRow: Integer;
    i: Integer;
begin
  currRow := sgTable.Selection.Top;
  for i := currRow to sgTable.RowCount-2 do
  begin
    sgTable.Rows[i].Assign(sgTable.Rows[i+1]);
    sgTable.Cells[0, i] := IntToStr(i);
  end;
  sgTable.RowCount := sgTable.RowCount - 1;

  SetModified(true);
end;

function TTableFrame.Find(AText: String; AStartPos: TPoint; AOptions: TFindOptions; out ResultPosition: TPoint): Boolean;
var y: Integer;
    row: String;
    p: Integer;
    tmpSL: TStringList;
    ok: Boolean;
begin
  Result := false;

  if not (frMatchCase in AOptions) then
    AText := AnsiLowerCase(AText);

  tmpSL := TStringList.Create;
  try
    if frDown in AOptions then
    begin
      for y := AStartPos.Y to sgTable.RowCount-1 do
      begin
        tmpSL.Assign(sgTable.Rows[y]);
        tmpSL.Delete(0);

        row := StringReplace(tmpSL.Text, sLineBreak, #9, [rfReplaceAll]);
        if not (frMatchCase in AOptions) then
          row := AnsiLowerCase(row);

        p := pos(AText, row);
        while (p > 0) do
        begin
          ok := true;
          if frWholeWord in AOptions then
          begin
            if not (((p = 1) or ((p-1 > 1) and ((row[p-1] = ' ') or (row[p-1] = #9)))) and
               ((p+Length(AText) >= Length(row)) or ((p+Length(AText) <= Length(row)) and ((row[p+Length(AText)] = ' ') or (row[p+Length(AText)] = #9)))))
            then
              ok := false;
          end;
          
          if ok then
          begin
            tmpSL.Delimiter := #9;
            tmpSL.StrictDelimiter := true;
            tmpSL.DelimitedText := copy(row, 1, p);
            if (y = AStartPos.Y) then
            begin
              if tmpSL.Count > AStartPos.X then
              begin
                ResultPosition := Point(tmpSL.Count, y);
                Result := true;
                break;
              end;
            end
            else
            begin
              ResultPosition := Point(tmpSL.Count, y);
              Result := true;
              break;
            end;
          end;

          p := posEx(Atext, row, p+1)
        end;

        if Result then
          break;
      end;
    end
    else
    begin
      AText := ReverseString(AText);
      for y := AStartPos.Y downto 0 do
      begin
        tmpSL.Assign(sgTable.Rows[y]);
        tmpSL.Delete(0);

        row := ReverseString(StringReplace(tmpSL.Text, sLineBreak, #9, [rfReplaceAll]));
        if not (frMatchCase in AOptions) then
          row := AnsiLowerCase(row);

        p := pos(AText, row);
        while (p > 0) do
        begin
          ok := true;
          if frWholeWord in AOptions then
          begin
            if not (((p = 1) or ((p-1 > 1) and ((row[p-1] = ' ') or (row[p-1] = #9)))) and
               ((p+Length(AText) >= Length(row)) or ((p+Length(AText) <= Length(row)) and ((row[p+Length(AText)] = ' ') or (row[p+Length(AText)] = #9)))))
            then
              ok := false;
          end;

          if ok then
          begin
            tmpSL.Delimiter := #9;
            tmpSL.StrictDelimiter := true;
            tmpSL.DelimitedText := copy(row, 1, p);
            if (y = AStartPos.Y) then
            begin
              if (sgTable.ColCount - tmpSL.Count + 1) < AStartPos.X then
              begin
                ResultPosition := Point(sgTable.ColCount - tmpSL.Count + 1, y);
                Result := true;
                break;
              end;
            end
            else
            begin
              ResultPosition := Point(sgTable.ColCount - tmpSL.Count + 1, y);
              Result := true;
              break;
            end;
          end;

          p := posEx(Atext, row, p+1)
        end;

        if Result then
          break;
      end;
    end;
  finally
    tmpSL.Free;
  end;
end;

function TTableFrame.GetTabsheet: TTabsheet;
begin
  Result := TTabsheet(Parent);
end;

procedure TTableFrame.LoadFile(const AFilename: String);
var tabFile: TStringList;
    row: TArray<String>;
    i,j: Integer;
begin
  FFilename := AFilename;

  tabFile := TStringList.Create;
  try
    tabFile.LoadFromFile(AFilename, TEncoding.ANSI);
    if tabFile.Count > 0 then
    begin
      row := tabFile[0].Split([#9]);

      sgTable.RowCount := tabFile.Count;
      sgTable.ColCount := Length(row) + 1;
      for i := 0 to tabFile.Count-1 do
      begin
        row := tabFile[i].Split([#9]);
        if i > 0 then
        begin
          sgTable.Cells[0, i] := IntToStr(i);
          sgTable.ColWidths[0] := Max(sgTable.ColWidths[0], sgTable.Canvas.TextWidth(IntToStr(i) + '    '));
        end;

        for j := 0 to High(row) do
        begin
          sgTable.Cells[j+1, i] := row[j];
          sgTable.ColWidths[j+1] := Max(sgTable.ColWidths[j+1], sgTable.Canvas.TextWidth(row[j] + '    '));
        end;
      end;
    end;
  finally
    tabFile.Free;
  end;
end;

procedure TTableFrame.miGridAddCopyClick(Sender: TObject);
var currRow: Integer;
begin
  currRow := sgTable.Selection.Top;
  sgTable.RowCount := sgTable.RowCount + 1;

  sgTable.Rows[sgTable.RowCount-1].Assign(sgTable.Rows[currRow]);
  sgTable.Cells[0, sgTable.RowCount-1] := IntToStr(sgTable.RowCount - 1);

  SetModified(true);
end;

procedure TTableFrame.miGridAddNewClick(Sender: TObject);
begin
  sgTable.RowCount := sgTable.RowCount + 1;
  sgTable.Cells[0, sgTable.RowCount-1] := IntToStr(sgTable.RowCount - 1);

  SetModified(true);
end;

procedure TTableFrame.miGridDeleteRowClick(Sender: TObject);
begin
  DeleteCurrentRow;
end;

procedure TTableFrame.miGridInsertCopyClick(Sender: TObject);
var currRow: Integer;
    i: Integer;
begin
  currRow := sgTable.Selection.Top;

  sgTable.RowCount := sgTable.RowCount + 1;
  for i := sgTable.RowCount-1 downto currRow+1 do
  begin
    sgTable.Rows[i].Assign(sgTable.Rows[i-1]);
    sgTable.Cells[0, i] := IntToStr(i);
  end;

  sgTable.Rows[currRow].Assign(sgTable.Rows[currRow+1]);
  sgTable.Cells[0, currRow] := IntToStr(currRow);

  SetModified(true);
end;

procedure TTableFrame.miGridInsertNewClick(Sender: TObject);
var currRow: Integer;
    i: Integer;
begin
  currRow := sgTable.Selection.Top;

  sgTable.RowCount := sgTable.RowCount + 1;
  for i := sgTable.RowCount-1 downto currRow+1 do
  begin
    sgTable.Rows[i].Assign(sgTable.Rows[i-1]);
    sgTable.Cells[0, i] := IntToStr(i);
  end;

  sgTable.Rows[currRow].Clear;
  sgTable.Cells[0, currRow] := IntToStr(currRow);

  SetModified(true);
end;

procedure TTableFrame.SaveFile;
var tabFile: TStringList;
    line: String;
    y,x: Integer;
begin
  tabFile := TStringList.Create;
  try
    for y := 0 to sgTable.RowCount-1 do
    begin
      line := '';
      for x := 1 to sgTable.ColCount-1 do
        line := line + sgTable.Cells[x,y] + #9;
      tabFile.Add(Trim(line));
    end;

    tabFile.SaveToFile(FFilename, TEncoding.ANSI);
    SetModified(false);
  finally
    tabFile.Free;
  end;
end;

procedure TTableFrame.Select(ACol, ARow: Integer; AScrollTo: Boolean);
var sel: TGridRect;
    tmpLen: Integer;
    fixedWidth: Integer;
    visibleRowCount: Integer;
    i: Integer;
begin
  sel.Left := ACol;
  sel.Right := ACol;
  sel.Top := ARow;
  sel.Bottom := ARow;
  sgTable.Selection := sel;

  if AScrollTo then
  begin
    fixedWidth := 0;
    for i := 0 to sgTable.FixedCols-1 do
      fixedWidth := fixedWidth + sgTable.ColWidths[i];

    tmpLen := sgTable.ColWidths[ACol];
    while (tmpLen < (sgTable.Width-fixedWidth)) and (ACol >= sgTable.FixedCols) do
    begin
      dec(ACol);
      tmpLen := tmpLen + sgTable.ColWidths[ACol];
    end;
    inc(ACol);
    sgTable.LeftCol := Max(ACol, sgTable.FixedCols);

    visibleRowCount := (sgTable.ClientRect.Height div sgTable.RowHeights[0]);
    if sgTable.RowCount > visibleRowCount then
      visibleRowCount := ((sgTable.ClientRect.Height - 18) div sgTable.RowHeights[0]);

    if ARow < sgTable.TopRow then
      sgTable.TopRow := ARow
    else
    if ARow >= (sgTable.TopRow + visibleRowCount - sgTable.FixedRows) then
      sgTable.TopRow := ARow - visibleRowCount + sgTable.FixedRows + 1;
  end;
end;

procedure TTableFrame.SetModified(AModified: Boolean);
begin
  if AModified <> FModified then
  begin
    FModified := AModified;
    if Modified then
      Tabsheet.Caption := Tabsheet.Caption + '*'
    else
      Tabsheet.Caption := String(Tabsheet.Caption).Trim(['*']);

    if Assigned(OnModifiedChanged) then
      OnModifiedChanged(Tabsheet);
  end;
end;

procedure TTableFrame.sgTableContextPopup(Sender: TObject; MousePos: TPoint;
  var Handled: Boolean);
var row, col: Integer;
begin
  sgTable.MouseToCell(MousePos.X, MousePos.Y, col, row);
  Select(col, row);

  miGridAddCopy.Enabled := (col <> -1) and (row <> -1);
  miGridInsertCopy.Enabled := (col <> -1) and (row <> -1);
  miGridDeleteRow.Enabled := (col <> -1) and (row <> -1);
end;

procedure TTableFrame.sgTableDrawCell(Sender: TObject; ACol, ARow: Integer;
  Rect: TRect; State: TGridDrawState);
begin
  if (not (gdFixed in State)) and (not (gdSelected in State)) and (ARow mod 2 = 0) then
  begin
    Rect.Left := Rect.Left - 4;

    sgTable.Canvas.Brush.Color := RGB(250,250,250);
    sgTable.Canvas.Font.Color := clWindowText;

    sgTable.Canvas.FillRect(Rect);
    sgTable.Canvas.TextRect(Rect, Rect.Left + 6, Rect.Top + 2, sgTable.Cells[ACol, ARow]);
  end;
end;

procedure TTableFrame.sgTableGetEditText(Sender: TObject; ACol, ARow: Integer;
  var Value: string);
begin
  FOldValue := Value;
end;

procedure TTableFrame.sgTableKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var num: Integer;
    value: String;
begin
  if (Key = VK_DELETE) and (ssCtrl in Shift) then
    DeleteCurrentRow
  else
  if (Key = VK_DELETE) and (Shift = []) and (sgTable.Cells[sgTable.Selection.Left, sgTable.Selection.Top] <> '') then
  begin
    sgTable.Cells[sgTable.Selection.Left, sgTable.Selection.Top] := '';
    SetModified(true);
  end
  else
  if ((Key = VK_OEM_PLUS) or (Key = VK_ADD)) and (ssCtrl in Shift) then
  begin
    if TryStrToInt(sgTable.Cells[sgTable.Selection.Left, sgTable.Selection.Top], num) then
    begin
      sgTable.Cells[sgTable.Selection.Left, sgTable.Selection.Top] := IntToStr(num+1);
      SetModified(true);
    end;
  end
  else
  if ((Key = VK_OEM_MINUS) or (Key = VK_SUBTRACT)) and (ssCtrl in Shift) then
  begin
    if TryStrToInt(sgTable.Cells[sgTable.Selection.Left, sgTable.Selection.Top], num) then
    begin
      sgTable.Cells[sgTable.Selection.Left, sgTable.Selection.Top] := IntToStr(num-1);
      SetModified(true);
    end;
  end
  else
  if (Key = VK_ESCAPE) and (sgTable.EditorMode) then
  begin
    sgTable.Cells[sgTable.Selection.Left, sgTable.Selection.Top] := FOldValue;
    sgTable.EditorMode := false;
  end
  else
  if (Key = VK_RETURN) and (sgTable.EditorMode) then
  begin
    value := sgTable.Cells[sgTable.Selection.Left, sgTable.Selection.Top];
    if Value <> FOldValue then
      SetModified(true);
  end
  else
  if (Key = Word(VkKeyScan('c'))) and (ssCtrl in Shift) then
    Clipboard.AsText := sgTable.Cells[sgTable.Selection.Left, sgTable.Selection.Top]
  else
  if (Key = Word(VkKeyScan('x'))) and (ssCtrl in Shift) then
  begin
    Clipboard.AsText := sgTable.Cells[sgTable.Selection.Left, sgTable.Selection.Top];
    sgTable.Cells[sgTable.Selection.Left, sgTable.Selection.Top] := '';
    SetModified(true);
  end
  else
  if (Key = Word(VkKeyScan('v'))) and (ssCtrl in Shift) then
  begin
    sgTable.Cells[sgTable.Selection.Left, sgTable.Selection.Top] := Clipboard.AsText;
    SetModified(true);
  end;
end;

procedure TTableFrame.sgTableMouseWheelDown(Sender: TObject; Shift: TShiftState;
  MousePos: TPoint; var Handled: Boolean);
var rowDisplayCount: Integer;
begin
  if ssShift in Shift then
    sgTable.LeftCol := sgTable.LeftCol + 1
  else
  begin
    rowDisplayCount := (sgTable.Height div sgTable.RowHeights[0]);
    if sgTable.RowCount > rowDisplayCount then
      rowDisplayCount := ((sgTable.Height - 18) div sgTable.RowHeights[0]);

    sgTable.TopRow := Min(sgTable.TopRow + 1, Max(sgTable.RowCount - rowDisplayCount + 3, 1));
  end;

  Handled := true;
end;

procedure TTableFrame.sgTableMouseWheelUp(Sender: TObject; Shift: TShiftState;
  MousePos: TPoint; var Handled: Boolean);
begin
  if ssShift in Shift then
    sgTable.LeftCol := Max(sgTable.LeftCol - 1, sgTable.FixedCols)
  else
    sgTable.TopRow := Max(sgTable.TopRow - 1, sgTable.FixedRows);

  Handled := true;
end;

procedure TTableFrame.sgTableSelectCell(Sender: TObject; ACol, ARow: Integer;
  var CanSelect: Boolean);
var value: String;
begin
  if sgTable.EditorMode then
  begin
    value := sgTable.Cells[sgTable.Selection.Left, sgTable.Selection.Top];
    if Value <> FOldValue then
      SetModified(true);
  end;
end;

end.
