program ArduinoComm;

uses
  Vcl.Forms,
  FrmArduinoComm in 'FrmArduinoComm.pas' {Form1},
  UnitSerialCommWorker in '..\..\..\common\UnitSerialCommWorker.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  ReportMemoryLeaksOnShutdown := DebugHook <> 0;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
