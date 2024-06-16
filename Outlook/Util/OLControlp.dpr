program OLControlp;

uses
  Vcl.Forms,
  FrmOLControl in 'FrmOLControl.pas' {OLControlF},
  UnitOLControlWorker in 'UnitOLControlWorker.pas',
  UnitOutLookDataType in '..\common\UnitOutLookDataType.pas',
  UnitSynLog2 in '..\..\..\Common\UnitSynLog2.pas',
  FrameOLEmailList4Ole in '..\..\..\Common\Frame\FrameOLEmailList4Ole.pas' {OutlookEmailListFr: TFrame},
  Outlook_TLB in '..\common\tlb\Outlook_TLB.pas',
  UnitOLEmailRecord2 in '..\..\..\Common\UnitOLEmailRecord2.pas',
  UnitElecServiceData2 in '..\..\GSManage\UnitElecServiceData2.pas';

{$R *.res}

begin
  ReportMemoryLeaksOnShutdown := DebugHook <> 0;
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TOLControlF, OLControlF);
  Application.Run;
end.
