unit uMain;

{$WARN SYMBOL_PLATFORM OFF}

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.Grids, Vcl.ExtCtrls, Vcl.ComCtrls,
  Vcl.Menus, Vcl.ExtDlgs, Generics.Collections, System.IOUtils, IniFiles, ShellAPI,
  Vcl.ToolWin, System.ImageList, Vcl.ImgList, Vcl.StdCtrls,
  uTableFrame, uTableFrameTabsheet;

type
  TExcelPlusMainForm = class(TForm)
    {$region 'Components'}
    plSideBar: TPanel;
    pcFiles: TPageControl;
    mmMain: TMainMenu;
    miFile: TMenuItem;
    miOpen: TMenuItem;
    miExit: TMenuItem;
    odOpenDialog: TFileOpenDialog;
    tbToolbar: TToolBar;
    tbOpen: TToolButton;
    tbSave: TToolButton;
    tbSeperator2: TToolButton;
    miSave: TMenuItem;
    imlIcons32: TImageList;
    tbSaveAll: TToolButton;
    tbSeperator1: TToolButton;
    miSaveAll: TMenuItem;
    pmFiles: TPopupMenu;
    miCloseFile: TMenuItem;
    tvWorkspace: TTreeView;
    odFolder: TFileOpenDialog;
    tbOpenFolder: TToolButton;
    miOpenWS: TMenuItem;
    imlIcons16: TImageList;
    miClose: TMenuItem;
    N1: TMenuItem;
    N3: TMenuItem;
    miCloseAll: TMenuItem;
    miCloseAllFiles: TMenuItem;
    spSplitter: TSplitter;
    N2: TMenuItem;
    miCloseOtherFiles: TMenuItem;
    fdSearch: TFindDialog;
    Edit1: TMenuItem;
    miFind: TMenuItem;
    tbSeperator3: TToolButton;
    tbSearch: TToolButton;
    miHelp: TMenuItem;
    miAbout: TMenuItem;
    {$endregion}
    {$region 'Eventhandler'}
    procedure miOpenClick(Sender: TObject);
    procedure miExitClick(Sender: TObject);
    procedure miCloseFileClick(Sender: TObject);
    procedure miSaveClick(Sender: TObject);
    procedure pcFilesChange(Sender: TObject);
    procedure miOpenWSClick(Sender: TObject);
    procedure tvWorkspaceCollapsing(Sender: TObject; Node: TTreeNode;
      var AllowCollapse: Boolean);
    procedure tvWorkspaceDblClick(Sender: TObject);
    procedure pcFilesContextPopup(Sender: TObject; MousePos: TPoint;
      var Handled: Boolean);
    procedure miSaveAllClick(Sender: TObject);
    procedure pcFilesMouseMove(Sender: TObject; Shift: TShiftState; X, Y: Integer);
    procedure pcFilesMouseLeave(Sender: TObject);
    procedure miCloseAllFilesClick(Sender: TObject);
    procedure miCloseOtherFilesClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure TabModifiedChanged(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure miCloseClick(Sender: TObject);
    procedure miFindClick(Sender: TObject);
    procedure fdSearchFind(Sender: TObject);
    procedure miAboutClick(Sender: TObject);
    {$endregion}
  private
    FWorkspacePath: String;
    FLastFindPos: TPoint;
    FFiles: TDictionary<string, TTableFrameTabsheet>;

    procedure SetControls(ASetFocus: Boolean = true);

    procedure LoadSettings;
    procedure SaveSettings;

    function IsAnyModified: Boolean;

    procedure LoadWorkspace(AWorkspacePath: String);
    procedure OpenFile(AFilename: String);
    procedure CloseFile(AFilename: String);
  protected
    procedure WMDropFiles(var Msg: TMessage); message WM_DROPFILES;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
  end;

var
  ExcelPlusMainForm: TExcelPlusMainForm;

implementation

{$R *.dfm}

uses
  System.UITypes, System.Types, uAbout;

procedure TExcelPlusMainForm.CloseFile(AFilename: String);
var tab: TTableFrameTabsheet;
    dlgResult: Integer;
begin
  if FFiles.ContainsKey(AFilename) then
  begin
    tab := FFiles[AFilename];
    if tab.Frame.Modified then
    begin
      dlgResult := MessageDlg('Do you want to save your changes to "' + ExtractFileName(AFilename) + '"?', mtWarning, [mbYes, mbNo, mbCancel], 0);
      if dlgResult <> mrCancel then
      begin
        if dlgResult = mrYes then
          tab.Frame.SaveFile();

        FFiles[AFilename].Free;
        FFiles.Remove(AFilename);
      end;
    end
    else
    begin
      FFiles[AFilename].Free;
      FFiles.Remove(AFilename);
    end;

    SetControls;
  end;
end;

constructor TExcelPlusMainForm.Create(AOwner: TComponent);
begin
  inherited;
  FFiles := TDictionary<String, TTableFrameTabsheet>.Create;
  DragAcceptFiles(Handle, true);
end;

destructor TExcelPlusMainForm.Destroy;
begin
  DragAcceptFiles(Handle, false);
  FreeAndNil(FFiles);
  inherited;
end;

procedure TExcelPlusMainForm.fdSearchFind(Sender: TObject);
var tab: TTableFrameTabsheet;
begin
  if frFindNext in fdSearch.Options then
  begin
    tab := TTableFrameTabsheet(pcFiles.ActivePage);
    if Assigned(tab) then
    begin
      FLastFindPos := Point(tab.Frame.sgTable.Selection.Left, tab.Frame.sgTable.Selection.Top);
      if tab.Frame.Find(fdSearch.FindText, FLastFindPos, fdSearch.Options, FLastFindPos) then
        tab.Frame.Select(FLastFindPos.X, FLastFindPos.Y, true)
      else
      begin
        if (frDown in fdSearch.Options) then
          FLastFindPos := Point(0,0)
        else
          FLastFindPos := Point(tab.Frame.sgTable.ColCount, tab.Frame.sgTable.RowCount);

        if tab.Frame.Find(fdSearch.FindText, FLastFindPos, fdSearch.Options, FLastFindPos) then
          tab.Frame.Select(FLastFindPos.X, FLastFindPos.Y, true)
        else
          MessageDlg('No results found!', mtInformation, [mbOK], 0);
      end
    end;
  end;
end;

procedure TExcelPlusMainForm.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  SaveSettings;
end;

procedure TExcelPlusMainForm.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
var msgResult: Integer;
    tab: TTableFrameTabsheet;
begin
  if IsAnyModified() then
  begin
    msgResult := MessageDlg('Do you want to save your changes before quitting?', mtWarning, [mbYes, mbNo, mbCancel], 0, mbCancel);
    if msgResult = mrYes then
    begin
      for tab in FFiles.Values do
      begin
        if tab.Frame.Modified then
          tab.Frame.SaveFile();
      end;
    end;

    CanClose := (msgResult <> mrCancel);
  end
  else
    CanClose := true;
end;

procedure TExcelPlusMainForm.FormCreate(Sender: TObject);
begin
  LoadSettings;
end;

function TExcelPlusMainForm.IsAnyModified: Boolean;
var tab: TTableFrameTabsheet;
begin
  Result := false;
  for tab in FFiles.Values do
  begin
    if tab.Frame.Modified then
    begin
      Result := true;
      break;
    end;
  end;
end;

procedure TExcelPlusMainForm.miAboutClick(Sender: TObject);
var aboutDlg: TAboutDlg;
begin
  aboutDlg := TAboutDlg.Create(Self);
  try
    aboutDlg.ShowModal;
  finally
    aboutDlg.Free;
  end;
end;

procedure TExcelPlusMainForm.miCloseAllFilesClick(Sender: TObject);
var i: Integer;
begin
  for i := pcFiles.PageCount-1 downto 0 do
    CloseFile(TTableFrameTabsheet(pcFiles.Pages[i]).Frame.Filename);
end;

procedure TExcelPlusMainForm.miCloseClick(Sender: TObject);
var tab: TTableFrameTabsheet;
begin
  tab := TTableFrameTabsheet(pcFiles.ActivePage);
  if Assigned(tab) then
    CloseFile(tab.Frame.Filename);
end;

procedure TExcelPlusMainForm.miCloseFileClick(Sender: TObject);
var tabIndex: Integer;
    tab: TTableFrameTabsheet;
    pt: TPoint;
begin
  pt := pcFiles.ScreenToClient(pmFiles.PopupPoint);
  tabIndex := pcFiles.IndexOfTabAt(pt.X, pt.Y);
  if tabIndex > -1 then
  begin
    tab := TTableFrameTabsheet(pcFiles.Pages[tabIndex]);
    CloseFile(tab.Frame.Filename);
  end;
end;

procedure TExcelPlusMainForm.miCloseOtherFilesClick(Sender: TObject);
var i, tabIndex: Integer;
    currPage: TTabsheet;
    pt: TPoint;
begin
  pt := pcFiles.ScreenToClient(pmFiles.PopupPoint);
  tabIndex := pcFiles.IndexOfTabAt(pt.X, pt.Y);
  if tabIndex > -1 then
  begin
    currPage := pcFiles.Pages[tabIndex];
    for i := pcFiles.PageCount-1 downto 0 do
    begin
      if pcFiles.Pages[i] <> currPage then
        CloseFile(TTableFrameTabsheet(pcFiles.Pages[i]).Frame.Filename);
    end;
  end;
end;

procedure TExcelPlusMainForm.miExitClick(Sender: TObject);
begin
  Close;
end;

procedure TExcelPlusMainForm.miFindClick(Sender: TObject);
var tab: TTableFrameTabsheet;
begin
  tab := TTableFrameTabsheet(pcFiles.ActivePage);
  if Assigned(tab) then
  begin
    FLastFindPos := Point(tab.Frame.sgTable.Selection.Left, tab.Frame.sgTable.Selection.Top);
    fdSearch.Execute();
  end;
end;

procedure TExcelPlusMainForm.miOpenClick(Sender: TObject);
var i: Integer;
begin
  if odOpenDialog.Execute() then
  begin
    for i := 0 to odOpenDialog.Files.Count-1 do
      OpenFile(odOpenDialog.Files[i]);
  end;
end;

procedure TExcelPlusMainForm.miOpenWSClick(Sender: TObject);
begin
  if odFolder.Execute then
    LoadWorkspace(odFolder.FileName);
end;

procedure TExcelPlusMainForm.miSaveAllClick(Sender: TObject);
var i: Integer;
    tab: TTableFrameTabsheet;
begin
  for i := 0 to pcFiles.PageCount-1 do
  begin
    tab := TTableFrameTabsheet(pcFiles.Pages[i]);
    tab.Frame.SaveFile();
  end;

  SetControls;
end;

procedure TExcelPlusMainForm.miSaveClick(Sender: TObject);
var tab: TTableFrameTabsheet;
begin
  if Assigned(pcFiles.ActivePage) then
  begin
    tab := TTableFrameTabsheet(pcFiles.ActivePage);
    tab.Frame.SaveFile();
  end;
end;

procedure TExcelPlusMainForm.OpenFile(AFilename: String);
var tabsheet: TTableFrameTabsheet;
begin
  if FFiles.ContainsKey(AFilename) then
    pcFiles.ActivePage := FFiles[AFilename]
  else
  begin
    tabsheet := TTableFrameTabsheet.Create(Self);
    tabsheet.PageControl := pcFiles;
    tabsheet.Caption := ExtractFileName(AFilename);
    tabsheet.Hint := AFilename;
    tabsheet.ParentShowHint := false;
    tabsheet.ShowHint := false;
    tabsheet.Frame := TTableFrame.Create(Self, AFilename);
    tabsheet.Frame.OnModifiedChanged := TabModifiedChanged;
    pcFiles.ActivePage := tabsheet;

    FFiles.Add(AFilename, tabsheet);
  end;

  SetControls;
end;

procedure TExcelPlusMainForm.pcFilesChange(Sender: TObject);
begin
  SetControls;
end;

procedure TExcelPlusMainForm.pcFilesContextPopup(Sender: TObject;
  MousePos: TPoint; var Handled: Boolean);
begin
  if not (htOnItem in pcFiles.GetHitTestInfoAt(MousePos.X, MousePos.Y)) then
    Handled := true;
end;

procedure TExcelPlusMainForm.pcFilesMouseLeave(Sender: TObject);
begin
  pcFiles.Hint := '';
end;

procedure TExcelPlusMainForm.pcFilesMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
var tabindex: integer;
    ht: THitTests;
begin
  tabindex := pcFiles.IndexOfTabAt(X, Y);
  if (tabindex >= 0) and (pcFiles.Hint <> pcFiles.Pages[tabindex].Hint) then
  begin
    Application.CancelHint;
    pcFiles.Hint := pcFiles.Pages[tabindex].Hint;
    ht := pcFiles.GetHitTestInfoAt(X, Y);
    pcFiles.ShowHint := htOnItem in ht;
  end;
end;

procedure TExcelPlusMainForm.LoadSettings;
var path: String;
    ini: TIniFile;
begin
  path := IncludeTrailingPathDelimiter(GetEnvironmentVariable('APPDATA'));
  path := path + 'D2ExcelPlus\';
  if FileExists(path + 'settings.ini') then
  begin
    ini := TIniFile.Create(path + 'settings.ini');
    try
      FWorkspacePath := ini.ReadString('SETTINGS', 'workspace', '');
      LoadWorkSpace(FWorkspacePath);
    finally
      ini.Free;
    end;
  end;
end;

procedure TExcelPlusMainForm.LoadWorkspace(AWorkspacePath: String);
var wsNode: TTreeNode;
    sr: TSearchRec;
begin
  FWorkspacePath := AWorkspacePath;

  tvWorkspace.Items.Clear;
  wsNode := tvWorkspace.Items.AddChild(nil, 'Workspace');
  wsNode.ImageIndex := 1;
  wsNode.ExpandedImageIndex := 1;
  wsNode.OverlayIndex := 1;
  wsNode.StateIndex := 1;
  wsNode.SelectedIndex := 1;

  try
    if FindFirst(IncludeTrailingPathDelimiter(FWorkspacePath) + '*.txt', faAnyFile, sr) = 0 then
    begin
      repeat
        tvWorkspace.Items.AddChild(wsNode, sr.Name);
      until FindNext(sr) <> 0;
    end;

    wsNode.Expand(true);
  finally
    FindClose(sr);
  end;
end;

procedure TExcelPlusMainForm.SaveSettings;
var path: String;
    ini: TIniFile;
begin
  path := IncludeTrailingPathDelimiter(GetEnvironmentVariable('APPDATA'));
  path := path + 'D2ExcelPlus\';
  ForceDirectories(path);
  ini := TIniFile.Create(path + 'settings.ini');
  try
    ini.WriteString('SETTINGS', 'workspace', FWorkspacePath);
    ini.UpdateFile;
  finally
    ini.Free;
  end;
end;

procedure TExcelPlusMainForm.SetControls(ASetFocus: Boolean);
var currTab: TTableFrameTabsheet;
    anyModified: Boolean;
begin
  currTab := TTableFrameTabsheet(pcFiles.ActivePage);
  if Assigned(currTab) then
  begin
    if ASetFocus then
      currTab.SetFocus;

    anyModified := IsAnyModified();

    Caption := ExtractFileName(currTab.Frame.Filename) + ' - D2Excel Plus';

    miSave.Enabled := currTab.Frame.Modified;
    tbSave.Enabled := currTab.Frame.Modified;
    miSaveAll.Enabled := anyModified;
    tbSaveAll.Enabled := anyModified;
    miClose.Enabled := true;
    miCloseAll.Enabled := true;
  end
  else
  begin
    Caption := 'D2Excel Plus';

    miClose.Enabled := false;
    miCloseAll.Enabled := false;
    miSave.Enabled := false;
    tbSave.Enabled := false;
    miSaveAll.Enabled := false;
    tbSaveAll.Enabled := false;
  end;
end;

procedure TExcelPlusMainForm.TabModifiedChanged(Sender: TObject);
begin
  if Sender = pcFiles.ActivePage then
    SetControls(false);
end;

procedure TExcelPlusMainForm.tvWorkspaceCollapsing(Sender: TObject;
  Node: TTreeNode; var AllowCollapse: Boolean);
begin
  AllowCollapse := Assigned(Node.Parent);
end;

procedure TExcelPlusMainForm.tvWorkspaceDblClick(Sender: TObject);
var selNode: TTreeNode;
begin
  selNode := tvWorkspace.Selected;
  if Assigned(selNode) and Assigned(selNode.Parent) then
    OpenFile(IncludeTrailingPathDelimiter(FWorkspacePath) + selNode.Text);
end;

procedure TExcelPlusMainForm.WMDropFiles(var Msg: TMessage);
var drop: THandle;
    fileCount: Integer;
    i, fileLen: Integer;
    filename: String;
begin
  drop := Msg.WParam;
  try
    fileCount := DragQueryFile(drop, $FFFFFFFF, nil, 0);
    for i := 0 to fileCount-1 do
    begin
      fileLen := DragQueryFile(drop, i, nil, 0) + 1;
      SetLength(filename, fileLen);
      DragQueryFile(drop, i, @filename[1], fileLen);
      OpenFile(Trim(filename));
    end;
  finally
    DragFinish(drop);
  end;
end;

end.
