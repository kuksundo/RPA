unit OLControlWorker;

interface

uses Windows, Winapi.Messages, System.SysUtils, System.SyncObjs, System.Classes,
  Variants, System.Win.ComObj,
  mormot.core.json,
  OtlComm, OtlCommon,
  UnitWorker4OmniMsgQ,
  UnitOutLookDataType;

type
  TOLControlWorker = class(TWorker2)
  strict private
    FOutlook: OLEVariant;

    procedure RespondEnqueueAndNotifyMainComm(AMsgId: word;
      const AValue: TOmniValue; const AWinMsg: integer); //MSG_ONTRCODE
  protected
    procedure ProcessCommandProc(AMsg: TOmniMessage); override;
    procedure ProcessRespondRealData(AMsg: TOmniMessage);
  public
    constructor Create(commandQueue, responseQueue, sendQueue: TOmniMessageQueue);
    destructor Destroy(); override;
    procedure Log2MainComm(const AMsg: string);
    procedure CustomCreate; override;
  end;

implementation

{ TOLControlWorker }

constructor TOLControlWorker.Create(commandQueue, responseQueue,
  sendQueue: TOmniMessageQueue);
begin
  inherited Create(commandQueue, responseQueue, sendQueue);
end;

procedure TOLControlWorker.CustomCreate;
begin
  if VarIsNull(FOutlook) then
  begin
    try
      FOutlook := GetActiveOleObject('outlook.application');
    except
      try
        FOutlook := CreateOleObject('outlook.application');
      except
        // Unable to access or start OUTLOOK
        Log2MainComm(
          'Unable to start or access Outlook.  Possibilities include: permission problems, server down, or VPN not enabled.  Exiting...', mtError, [mbOK], 0);
        exit;
      end;
    end;
  end;
end;

destructor TOLControlWorker.Destroy;
begin
  if not VarIsNull(FOutlook) then
  begin

    FOutlook := null;
  end;

  inherited;
end;

procedure TOLControlWorker.Log2MainComm(const AMsg: string);
begin
  RespondEnqueueAndNotifyMainComm(Ord(olrkLog), TOmniValue.CastFrom(AMsg), MSG_RESULT);
end;

procedure TOLControlWorker.ProcessCommandProc(AMsg: TOmniMessage);
var
  LValue: TOmniValue;
begin
  case TOLCommandKind(AMsg.MsgID) of
    olckAddAppointment: ;
  end;

end;

procedure TOLControlWorker.ProcessRespondRealData(AMsg: TOmniMessage);
begin

end;

procedure TOLControlWorker.RespondEnqueueAndNotifyMainComm(AMsgId: word;
  const AValue: TOmniValue; const AWinMsg: integer);
begin

end;

end.
