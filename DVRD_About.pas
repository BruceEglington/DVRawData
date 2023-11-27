unit DVRD_About;

interface

uses Windows, Classes, Graphics, Forms, Controls, StdCtrls,
  Buttons, ExtCtrls, System.ImageList, Vcl.ImgList, Vcl.VirtualImageList;

type
  TAboutBox = class(TForm)
    Panel1: TPanel;
    OKButton: TButton;
    ProgramIcon: TImage;
    ProductName: TLabel;
    Version: TLabel;
    Copyright: TLabel;
    Memo1: TMemo;
    Panel2: TPanel;
    Panel3: TPanel;
    VirtualImageList1: TVirtualImageList;
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  AboutBox: TAboutBox;

implementation

uses DVRD_varb, DVRD_dm;

{$R *.DFM}

procedure TAboutBox.FormShow(Sender: TObject);
begin
  Version.Caption := 'Version '+ DVRDVersion;
end;

end.

