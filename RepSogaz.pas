unit RepSogaz;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DBGridEh, StdCtrls, Mask, DBCtrlsEh, DBLookupEh, ComCtrls;

type
  TfRepSogaz = class(TForm)
    lbl1: TLabel;
    lbl2: TLabel;
    lbl3: TLabel;
    dtp1: TDateTimePicker;
    dtp2: TDateTimePicker;
    cbb1: TDBLookupComboboxEh;
    btn1: TButton;
    btn2: TButton;
    procedure FormCreate(Sender: TObject);
    procedure btn2Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fRepSogaz: TfRepSogaz;

implementation
uses
    DM,DateUtils;
{$R *.dfm}

procedure TfRepSogaz.btn2Click(Sender: TObject);
begin
//fRepSogazPat.ShowModal;
end;

procedure TfRepSogaz.FormCreate(Sender: TObject);
begin
dtp1.DateTime := StartOfTheMonth(Now);
dtp2.DateTime := EndOfTheMonth(Now);

end;

end.
