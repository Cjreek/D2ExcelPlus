program D2ExcelPlus;

uses
  Vcl.Forms,
  uMain in 'uMain.pas' {ExcelPlusMainForm},
  uTableFrame in 'uTableFrame.pas' {TableFrame: TFrame},
  uAbout in 'uAbout.pas' {AboutDlg},
  uTableFrameTabsheet in 'uTableFrameTabsheet.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TExcelPlusMainForm, ExcelPlusMainForm);
  Application.Run;
end.
