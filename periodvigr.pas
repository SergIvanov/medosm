unit periodvigr;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DBGridEhGrouping, GridsEh, DBGridEh, ExtCtrls;

type
  TfPeriodVigr = class(TForm)
    Panel1: TPanel;
    DBGridEh1: TDBGridEh;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fPeriodVigr: TfPeriodVigr;

implementation

uses DM;

{$R *.dfm}

end.
