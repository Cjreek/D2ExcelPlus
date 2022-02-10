unit uTableFrame;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.Grids, Vcl.ComCtrls,
  Vcl.ExtCtrls, Vcl.StdCtrls, Vcl.Samples.Spin, Vcl.Menus, Clipbrd, Generics.Collections;

type
  TEditorState = class
  strict private
    FArea: TGridREct;
    FValues: TStrings;
  public
    constructor Create;
    destructor Destroy; override;
    property Area: TGridREct read FArea write FArea;
    property Values: TStrings read FValues;
  end;

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
    procedure sgTableMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure sgTableMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure sgTableFixedCellClick(Sender: TObject; ACol, ARow: Integer);
    {$endregion}
  private
    FFilename: String;
    FOldValue: String;
    FModified: Boolean;
    FSelStart: TPoint;
    FOnModifiedChanged: TNotifyEvent;
    FUndoStack: TStack<TEditorState>;
    FRedoStack: TStack<TEditorState>;
    procedure DeleteCurrentRow;
    procedure SetModified(AModified: Boolean);
    function GetTabsheet: TTabsheet;
    function GetCanUndo: Boolean;
    function GetCanRedo: Boolean;

    procedure ClearStack(AStack: TStack<TEditorState>);
    function CreateEditorState(AArea: TGridRect): TEditorState;
    procedure ProcessEditorState(AFromStack: TStack<TEditorState>; AToStack: TStack<TEditorState>);
  public
    constructor Create(AOwner: TComponent; AFilename: String); reintroduce;
    destructor Destroy; override;

    procedure LoadFile(const AFilename: String);
    procedure SaveFile();

    function Find(AText: String; AStartPos: TPoint; AOptions: TFindOptions; out ResultPosition: TPoint): Boolean;
    procedure Select(ACol, ARow: Integer; AScrollTo: Boolean = false);

    procedure CreateUndo(); overload;
    procedure CreateUndo(AArea: TGridRect); overload;
    procedure Undo();
    procedure Redo();

    property Filename: String read FFilename;
    property Modified: Boolean read FModified;
    property CanUndo: Boolean read GetCanUndo;
    property CanRedo: Boolean read GetCanRedo;
    property Tabsheet: TTabsheet read GetTabsheet;

    property OnModifiedChanged: TNotifyEvent read FOnModifiedChanged write FOnModifiedChanged;
  end;

implementation

{$R *.dfm}

uses
  Math, StrUtils, System.Types, System.IOUtils, System.Hash;

const
  clAlternatingRow = $00FAFAFA;

{ TTableFrame }

procedure TTableFrame.cbFixColumnsClick(Sender: TObject);
var selection: TGridRect;
    leftCol, topRow: Integer;
begin
  seFixedColumns.Enabled := cbFixColumns.Checked;

  selection := sgTable.Selection;
  leftCol := sgTable.LeftCol;
  topRow := sgTable.TopRow;
  try
    if cbFixColumns.Checked then
    begin
      sgTable.FixedCols := seFixedColumns.Value + 1
    end
    else
    begin
      if leftCol = sgTable.FixedCols then
        leftCol := 1;
      sgTable.FixedCols := 1;
    end;
  finally
    selection.Left := Max(selection.Left, sgTable.FixedCols);
    selection.Right := Max(selection.Right, sgTable.FixedCols);
    sgTable.Selection := selection;

    sgTable.LeftCol := Max(leftCol, sgTable.FixedCols);
    sgTable.TopRow := topRow;
  end;
end;

procedure TTableFrame.ClearStack(AStack: TStack<TEditorState>);
var state: TEditorState;
begin
  while AStack.Count > 0 do
  begin
    state := AStack.Pop;
    state.Free;
  end;
end;

constructor TTableFrame.Create(AOwner: TComponent; AFilename: String);
begin
  inherited Create(AOwner);
  FUndoStack := TStack<TEditorState>.Create();
  FRedoStack := TStack<TEditorState>.Create();
  Name := 'TableFrame' + THashSHA2.GetHashString(AFilename);
  LoadFile(AFilename);
end;

function TTableFrame.CreateEditorState(AArea: TGridRect): TEditorState;
var x,y: Integer;
    line: String;
begin
  Result := TEditorState.Create;
  Result.Area := AArea;
  for y := Result.Area.Top to Result.Area.Bottom do
  begin
    line := '';
    for x := Result.Area.Left to Result.Area.Right do
      line := line + sgTable.Cells[x, y] + #9;
    Result.Values.Add(line)
  end;
end;

procedure TTableFrame.CreateUndo(AArea: TGridRect);
var state: TEditorState;
begin
  ClearStack(FRedoStack);

  state := CreateEditorState(AArea);
  FUndoStack.Push(state);
  if Assigned(FOnModifiedChanged) then
    FOnModifiedChanged(Tabsheet);
end;

procedure TTableFrame.CreateUndo;
begin
  CreateUndo(sgTable.Selection);
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

destructor TTableFrame.Destroy;
begin
  FreeAndNil(FUndoStack);
  FreeAndNil(FRedoStack);
  inherited;
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

function TTableFrame.GetCanRedo: Boolean;
begin
  Result := FRedoStack.Count > 0;
end;

function TTableFrame.GetCanUndo: Boolean;
begin
  Result := FUndoStack.Count > 0;
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

procedure TTableFrame.ProcessEditorState(AFromStack,
  AToStack: TStack<TEditorState>);
var state, newState: TEditorState;
    x,y: Integer;
    lineValues: TStringList;
begin
  if AFromStack.Count > 0 then
  begin
    state := AFromStack.Pop;
    try
      newState := CreateEditorState(state.Area);
      try
        lineValues := TStringList.Create;
        try
          lineValues.StrictDelimiter := true;
          lineValues.Delimiter := #9;
          for y := state.Area.Top to state.Area.Bottom do
          begin
            lineValues.DelimitedText := state.Values[y-state.Area.Top];
            for x := state.Area.Left to state.Area.Right do
              sgTable.Cells[x,y] := lineValues[x-state.Area.Left];
          end;
        finally
          lineValues.Free;
        end;
      finally
        AToStack.Push(newState);
      end;
    finally
      state.Free;
    end;

    if Assigned(FOnModifiedChanged) then
      FOnModifiedChanged(Tabsheet);
  end;
end;

procedure TTableFrame.Redo;
begin
  ProcessEditorState(FRedoStack, FUndoStack);
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
      SetLength(line, Length(line)-1);
      tabFile.Add(line);
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
  InflateRect(Rect, 1, 1);

  if gdFixed in State then
    sgTable.Canvas.Brush.Color := clBtnFace
  else
  if (ARow mod 2 = 0) then
  begin
    sgTable.Canvas.Brush.Color := clAlternatingRow;
    if (gdSelected in State) or (gdFocused in State) or (gdPressed in State) or (gdHotTrack in State) then
      sgTable.Canvas.Brush.Color := RGB(234, 239, 250);
  end
  else
  begin
    sgTable.Canvas.Brush.Color := clWindow;
    if (gdSelected in State) or (gdFocused in State) or (gdPressed in State) or (gdHotTrack in State) then
      sgTable.Canvas.Brush.Color := RGB(239, 243, 255);
  end;

  sgTable.Canvas.Font.Color := clWindowText;
  sgTable.Canvas.TextRect(Rect, Rect.Left + 6, Rect.Top + 2, sgTable.Cells[ACol, ARow]);

  sgTable.Canvas.Pen.Color := clSilver;
  sgTable.Canvas.Brush.Style := bsClear;
  sgTable.Canvas.Rectangle(Rect);

  InflateRect(Rect, -1, -1);

  if (ACol >= sgTable.Selection.Left) and (ACol <= sgTable.Selection.Right) and (ARow >= sgTable.Selection.Top) and (ARow <= sgTable.Selection.Bottom) then
  begin
    sgTable.Canvas.Pen.Color := RGB(0,128,192);
    if (ACol = sgTable.Selection.Left) then
    begin
      sgTable.Canvas.MoveTo(Rect.Left, Rect.Top);
      sgTable.Canvas.LineTo(Rect.Left, Rect.Bottom);
    end;

    if (ACol = sgTable.Selection.Right) then
    begin
      sgTable.Canvas.MoveTo(Rect.Right-1, Rect.Top);
      sgTable.Canvas.LineTo(Rect.Right-1, Rect.Bottom);
    end;

    if (ARow = sgTable.Selection.Top) then
    begin
      sgTable.Canvas.MoveTo(Rect.Left, Rect.Top);
      sgTable.Canvas.LineTo(Rect.Right, Rect.Top);
    end;

    if (ARow = sgTable.Selection.Bottom) then
    begin
      sgTable.Canvas.MoveTo(Rect.Left, Rect.Bottom-1);
      sgTable.Canvas.LineTo(Rect.Right, Rect.Bottom-1);
    end;
  end;
end;

procedure TTableFrame.sgTableFixedCellClick(Sender: TObject; ACol,
  ARow: Integer);
var gr: TGridRect;
begin
  if (ARow > 0) then
  begin
    gr.Left := sgTable.FixedCols;
    gr.Right := sgTable.ColCount-1;
    if GetAsyncKeyState(VK_SHIFT) < 0 then
    begin
      gr.Top := Min(ARow, sgTable.Selection.Top);
      gr.Bottom := Max(ARow, sgTable.Selection.Bottom);
    end
    else
    begin
      gr.Top := ARow;
      gr.Bottom := ARow;
    end;
    sgTable.Selection := gr;
  end
  else if (ACol >= sgTable.FixedCols) then
  begin
    if GetAsyncKeyState(VK_SHIFT) < 0 then
    begin
      gr.Left := Min(ACol, sgTable.Selection.Left);
      gr.Right := Max(ACol, sgTable.Selection.Right);
    end
    else
    begin
      gr.Left := ACol;
      gr.Right := ACol;
    end;
    gr.Top := sgTable.FixedRows;
    gr.Bottom := sgTable.RowCount-1;
    sgTable.Selection := gr;
  end;
end;

procedure TTableFrame.sgTableGetEditText(Sender: TObject; ACol, ARow: Integer;
  var Value: string);
begin
  if Value <> FOldValue then
    FOldValue := Value;
  CreateUndo();
end;

procedure TTableFrame.sgTableKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var num: Integer;
    value, line: String;
    sl: TStringList;
    rowValues: TArray<String>;
    x,y: Integer;
    gr: TGridRect;
begin
  if Key = VK_CONTROL then
    Key := 0
  else
  if (Key = VK_DELETE) and (ssCtrl in Shift) then
    DeleteCurrentRow
  else
  if (Key = VK_DELETE) and (Shift = []) and (sgTable.Cells[sgTable.Selection.Left, sgTable.Selection.Top] <> '') then
  begin
    CreateUndo();
    for y := sgTable.Selection.Top to sgTable.Selection.Bottom do
    begin
      for x := sgTable.Selection.Left to sgTable.Selection.Right do
        sgTable.Cells[x, y] := '';
    end;
    SetModified(true);
  end
  else
  if ((Key = VK_OEM_PLUS) or (Key = VK_ADD)) and (ssCtrl in Shift) then
  begin
    CreateUndo();
    for y := sgTable.Selection.Top to sgTable.Selection.Bottom do
    begin
      for x := sgTable.Selection.Left to sgTable.Selection.Right do
      begin
        if TryStrToInt(sgTable.Cells[x, y], num) then
        begin
          sgTable.Cells[x, y] := IntToStr(num+1);
          SetModified(true);
        end;
      end;
    end;
  end
  else
  if ((Key = VK_OEM_MINUS) or (Key = VK_SUBTRACT)) and (ssCtrl in Shift) then
  begin
    for y := sgTable.Selection.Top to sgTable.Selection.Bottom do
    begin
      for x := sgTable.Selection.Left to sgTable.Selection.Right do
      begin
        if TryStrToInt(sgTable.Cells[x, y], num) then
        begin
          sgTable.Cells[x, y] := IntToStr(num-1);
          SetModified(true);
        end;
      end;
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
  begin
    sl := TStringList.Create;
    try
      for y := sgTable.Selection.Top to sgTable.Selection.Bottom do
      begin
        line := '';
        for x := sgTable.Selection.Left to sgTable.Selection.Right do
          line := line + sgTable.Cells[x, y] + #9;
        sl.Add(line)
      end;
      Clipboard.AsText := sl.Text.Trim([#13, #10, #9]);
    finally
      sl.Free;
    end;
  end
  else
  if (Key = Word(VkKeyScan('x'))) and (ssCtrl in Shift) then
  begin
    CreateUndo;

    sl := TStringList.Create;
    try
      for y := sgTable.Selection.Top to sgTable.Selection.Bottom do
      begin
        line := '';
        for x := sgTable.Selection.Left to sgTable.Selection.Right do
        begin
          line := line + sgTable.Cells[x, y] + #9;
          sgTable.Cells[x, y] := '';
        end;
        sl.Add(line)
      end;
      Clipboard.AsText := sl.Text.Trim([#13, #10, #9]);
    finally
      sl.Free;
    end;

    SetModified(true);
  end
  else
  if (Key = Word(VkKeyScan('v'))) and (ssCtrl in Shift) then
  begin
    if Clipboard.HasFormat(CF_TEXT) then
    begin
      sl := TStringList.Create;
      try
        sl.Text := Clipboard.AsText;
        rowValues := sl[0].Split([#9]);

        gr.Left := sgTable.Selection.Left;
        gr.Top := sgTable.Selection.Top;
        gr.Right := gr.Left + Length(rowValues) - 1;
        gr.Bottom := gr.Top + sl.Count - 1;
        CreateUndo(gr);

        for y := sgTable.Selection.Top to Min(sgTable.Selection.Top + sl.Count-1, sgTable.RowCount-1)  do
        begin
          line := sl[y-sgTable.Selection.Top];
          rowValues := line.Split([#9]);
          for x := sgTable.Selection.Left to Min(sgTable.Selection.Left + High(rowValues), sgTable.ColCount-1) do
            sgTable.Cells[x, y] := rowValues[x-sgTable.Selection.Left];
          sl.Add(line)
        end;
      finally
        sl.Free;
      end;

      SetModified(true);
    end;
  end;
end;

procedure TTableFrame.sgTableMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
var row, col: Integer;
begin
  sgTable.MouseToCell(X, Y, col, row);
  FSelStart := Point(col, row);
end;

procedure TTableFrame.sgTableMouseMove(Sender: TObject; Shift: TShiftState; X,
  Y: Integer);
var col, row: Integer;
    gr: TGridRect;
    maxRow, maxCol: Boolean;
begin
  if (ssLeft in Shift) and (not sgTable.EditorMode) then
  begin
    maxRow := sgTable.Selection.Bottom = sgTable.RowCount-1;
    maxCol := sgTable.Selection.Right = sgTable.ColCount-1;

    sgTable.MouseToCell(X, Y, col, row);
    if (col = -1) and (row = -1) then
    begin
      if maxRow then
      begin
        repeat
          dec(Y);
          sgTable.MouseToCell(X, Y, col, row);
        until ((col <> -1) and (row <> -1)) or (Y < 0);
      end;

      if maxCol then
      begin
        repeat
          dec(X);
          sgTable.MouseToCell(X, Y, col, row);
        until ((col <> -1) and (row <> -1)) or (X < 0);
      end;
    end;

    if (X >= 0) and (Y >= 0) then
    begin
      gr.Left := Min(col, FSelStart.X);
      gr.Right := Max(col, FSelStart.X);
      gr.Top := Min(row, FSelStart.Y);
      gr.Bottom := Max(row, FSelStart.Y);

      sgTable.Selection := gr;
      sgTable.Repaint;
    end;
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
    newSelection: TGridRect;
begin
  if sgTable.EditorMode then
  begin
    value := sgTable.Cells[sgTable.Selection.Right, sgTable.Selection.Bottom];
    if Value <> FOldValue then
      SetModified(true)
    else
      FUndoStack.Pop.Free;
  end
  else if GetAsyncKeyState(VK_SHIFT) < 0 then
  begin
    newSelection.Left := sgTable.Selection.Left;
    newSelection.Top := sgTable.Selection.Top;
    newSelection.Right := ACol;
    newSelection.Bottom := ARow;
    sgTable.Selection := newSelection;
    sgTable.Repaint;
    CanSelect := false;
  end;
end;

procedure TTableFrame.Undo;
begin
  ProcessEditorState(FUndoStack, FRedoStack);
end;

{ TEditorState }

constructor TEditorState.Create;
begin
  FValues := TStringList.Create;
end;

destructor TEditorState.Destroy;
begin
  FreeAndNil(FValues);
  inherited;
end;

end.
