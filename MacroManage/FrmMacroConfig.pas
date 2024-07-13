unit FrmMacroConfig;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, JvExControls, JvLabel,
  Vcl.Buttons, Vcl.ExtCtrls,
  UnitMacroListClass2;

type
  TMacroConfigF = class(TForm)
    JvLabel6: TJvLabel;
    MacroName: TEdit;
    JvLabel1: TJvLabel;
    MacroDesc: TEdit;
    JvLabel2: TJvLabel;
    RepeatCount: TEdit;
    JvLabel3: TJvLabel;
    ActionDesc: TEdit;
    IsExecute: TCheckBox;
    IsDisplayCustomDesc: TCheckBox;
    Panel1: TPanel;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

  function ShowMacroConfig(var AMacroM: TMacroManagement): integer;

implementation

uses UnitIniConfigBase;

{$R *.dfm}

function ShowMacroConfig(var AMacroM: TMacroManagement): integer;
var
  MacroConfigF: TMacroConfigF;
begin
  MacroConfigF := TMacroConfigF.Create(nil);

  with MacroConfigF do
  begin
    TJHPIniConfigBase.LoadObject2Form(MacroConfigF, AMacroM, True, False);
    try
      Result := ShowModal;

      if Result = mrOK then
      begin
        TJHPIniConfigBase.LoadForm2Object(MacroConfigF, AMacroM, True, False);
      end;
    finally
      Free;
    end;
  end;
end;

end.
