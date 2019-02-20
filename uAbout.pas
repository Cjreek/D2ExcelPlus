unit uAbout;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls, ShellApi,
  Vcl.Imaging.pngimage;

type
  TAboutDlg = class(TForm)
    {$region 'Components'}
    lAboutLine1: TLabel;
    lAboutLine2: TLabel;
    lAboutLine3: TLabel;
    lAboutLine4: TLabel;
    llCjreek: TLinkLabel;
    imgLogo: TImage;
    lAboutLine5: TLabel;
    llModblackmoon: TLinkLabel;
    lAboutLine6: TLabel;
    {$endregion}
    {$region 'Eventhandler'}
    procedure llCjreekLinkClick(Sender: TObject; const Link: string;
      LinkType: TSysLinkType);
    {$endregion}
  private
    { Private-Deklarationen }
  public
    { Public-Deklarationen }
  end;

implementation

{$R *.dfm}

procedure TAboutDlg.llCjreekLinkClick(Sender: TObject; const Link: string;
  LinkType: TSysLinkType);
begin
  ShellExecute(Handle, 'open', PChar(Link), nil, nil, SW_NORMAL);
end;

end.
