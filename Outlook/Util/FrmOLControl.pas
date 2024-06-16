unit FrmOLControl;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls,
  AdvOfficePager,
  mormot.core.log, mormot.core.base,
  OtlComm, OtlCommon,
  UnitWorker4OmniMsgQ,
  UnitOutLookDataType, UnitOLControlWorker, FrameOLEmailList4Ole;

type
  TOLControlF = class(TForm)
    Panel1: TPanel;
    AdvOfficePager1: TAdvOfficePager;
    LogPage: TAdvOfficePage;
    Splitter1: TSplitter;
    Edit1: TEdit;
    Button1: TButton;
    Button2: TButton;
    Memo1: TMemo;
    Button3: TButton;
    AdvOfficePage1: TAdvOfficePage;
    TOutlookEmailListFr1: TOutlookEmailListFr;
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Button1Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
  private
    FOLControlWorker: TOLControlWorker;
    FCommandQueue    : TOmniMessageQueue;
    FResponseQueue   : TOmniMessageQueue;
    FSendMsgQueue    : TOmniMessageQueue;
  protected
    procedure InitVar();
    procedure DestroyVar();

    procedure StartWorker;
    procedure StopWorker;
    procedure OnWorkerResult(var Msg: TMessage); message MSG_RESULT;
    procedure SendCmd2WorkerThrd(const ACmd: TOLCommandKind; const AValue: TOmniValue);
  public
    procedure Log(AMsg: string; AMemo: TMemo=nil);
  end;

var
  OLControlF: TOLControlF;

implementation

{$R *.dfm}

uses UnitSynLog2;

{ TOLControlF }

procedure TOLControlF.Button1Click(Sender: TObject);
begin
  SendCmd2WorkerThrd(olckGetFolderList, TOmniValue.CastFrom(''));
end;

procedure TOLControlF.Button3Click(Sender: TObject);
begin
  SendCmd2WorkerThrd(olckInitVar, TOmniValue.CastFrom(''));
end;

procedure TOLControlF.DestroyVar;
begin
  StopWorker();
end;

procedure TOLControlF.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  DestroyVar();
end;

procedure TOLControlF.FormCreate(Sender: TObject);
begin
  InitVar();
end;

procedure TOLControlF.InitVar;
begin
  InitSynLog();
  StartWorker();
end;

procedure TOLControlF.Log(AMsg: string; AMemo: TMemo);
begin
  if AMemo = nil then
    AMemo := Memo1;

  if AMemo.Lines.Count > MEMO_LOG_MAX_LINE_COUNT then
    AMemo.Lines.Clear;

  AMemo.Lines.Add(AMsg);

  DoLog(AMsg, False, sllInfo);
end;

procedure TOLControlF.OnWorkerResult(var Msg: TMessage);
var
  LMsg  : TOmniMessage;
  LValue: TOmniValue;
  LOLRespondRec: TOLRespondRec;
begin
  if FResponseQueue.TryDequeue(LMsg) then
  begin
    LOLRespondRec := LMsg.MsgData.ToRecord<TOLRespondRec>;

    case TOLRespondKind(LMsg.MsgID) of
      olrkMAPIFolderList: begin
        Log(LOLRespondRec.FMsg);
      end;
      olrkLog: Log(LOLRespondRec.FMsg);
    end;
  end;
end;

procedure TOLControlF.SendCmd2WorkerThrd(const ACmd: TOLCommandKind;
  const AValue: TOmniValue);
begin
  if not FCommandQueue.Enqueue(TOmniMessage.Create(Ord(ACmd), AValue)) then
    raise Exception.Create('Command queue is full!');
end;

procedure TOLControlF.StartWorker;
begin
  FCommandQueue := TOmniMessageQueue.Create(1000);
  FResponseQueue := TOmniMessageQueue.Create(1000, false);
  FSendMsgQueue := TOmniMessageQueue.Create(1000);

  FOLControlWorker := TOLControlWorker.Create(FCommandQueue, FResponseQueue, FSendMsgQueue, Self.Handle);
//  FOLControlWorker.FormHandle := Self.Handle;
end;

procedure TOLControlF.StopWorker;
begin
  if Assigned(FOLControlWorker) then
  begin
    TWorker(FOLControlWorker).Stop;
    FOLControlWorker.WaitFor;
    FreeAndNil(FOLControlWorker);
  end;

  FCommandQueue.Free;
  FResponseQueue.Free;
  FSendMsgQueue.Free;
end;

end.


