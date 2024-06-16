unit FrmArduinoComm;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, CPortCtl, Vcl.Buttons,
  OtlComm, OtlCommon,
  UnitSerialCommWorker;

type
  TForm1 = class(TForm)
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    BitBtn3: TBitBtn;
    BitBtn4: TBitBtn;
    ComComboBox1: TComComboBox;
    Label1: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BitBtn1Click(Sender: TObject);
  private
    FSerialCommWorker: TSerialCommWorker;
    FCommandQueue : TOmniMessageQueue;
    FResponseQueue: TOmniMessageQueue;
  public
    procedure SendCommand2Worker(const ACmd: TCommCommandType; const AValue: TOmniValue);
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.BitBtn1Click(Sender: TObject);
var
  LValue: TOmniValue;
  LSerialCommCmdRec: TSerialCommCmdRec;
begin
  LValue := TOmniValue.FromRecord(LSerialCommCmdRec);
  SendCommand2Worker(cctConnect, LValue);
end;

procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  if Assigned(FSerialCommWorker) then
  begin
    FSerialCommWorker.Stop;
    FSerialCommWorker.WaitFor;
    FreeAndNil(FSerialCommWorker);
  end;

  FCommandQueue.Free;
  FResponseQueue.Free;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  FCommandQueue := TOmniMessageQueue.Create(1000);
  FResponseQueue := TOmniMessageQueue.Create(1000);

  FSerialCommWorker := TSerialCommWorker.Create(FCommandQueue, FResponseQueue, nil);
  FSerialCommWorker.SetMainFormHandle(Self.Handle);
end;

procedure TForm1.SendCommand2Worker(const ACmd: TCommCommandType; const AValue: TOmniValue);
begin
  if not FCommandQueue.Enqueue(TOmniMessage.Create(Ord(ACmd), AValue)) then
    raise Exception.Create('Command queue is full!');
end;

end.
