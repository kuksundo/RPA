program MacroManagement2;

uses
  Vcl.Forms,
  UnitMacroRecorderMain2 in 'UnitMacroRecorderMain2.pas' {MacroManageF},
  UnitAction2 in 'UnitAction2.pas' {frmActions},
  UnitMacroListClass2 in 'UnitMacroListClass2.pas',
  UnitMacroConfigClass2 in 'UnitMacroConfigClass2.pas',
  FrmEventCaptureConfig in 'FrmEventCaptureConfig.pas' {EventCaptureConfigF};

{$R *.res}

begin
  ReportMemoryLeaksOnShutdown := DebugHook <> 0;
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TMacroManageF, MacroManageF);
  Application.Run;
end.
