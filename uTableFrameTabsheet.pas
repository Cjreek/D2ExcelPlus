unit uTableFrameTabsheet;

interface

uses
  Vcl.Controls, Vcl.ComCtrls, uTableFrame;

type
  TTableFrameTabsheet = class(TTabsheet)
  strict private
    FFrame: TTableFrame;
  private
    procedure SetFrame(const Value: TTableFrame);
  public
    procedure SetFocus; override;
    property Frame: TTableFrame read FFrame write SetFrame;
  end;

implementation

{ TTableFrameTabsheet }

procedure TTableFrameTabsheet.SetFocus;
begin
  inherited;
  if Assigned(FFrame) then
    Frame.sgTable.SetFocus;
end;

procedure TTableFrameTabsheet.SetFrame(const Value: TTableFrame);
begin
  FFrame := Value;
  FFrame.Parent := Self;
  FFrame.Align := alClient;
end;

end.
