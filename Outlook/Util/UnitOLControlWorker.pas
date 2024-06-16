unit UnitOLControlWorker;

interface

uses Windows, Winapi.Messages, System.SysUtils, System.SyncObjs, System.Classes,
  Variants, System.Win.ComObj, Vcl.ComCtrls, ActiveX,
  mormot.core.json, mormot.core.data, mormot.core.base, mormot.core.variants,
  OtlComm, OtlCommon,
  UnitWorker4OmniMsgQ, Outlook_TLB,
  UnitOutLookDataType;

type
  TOLControlWorker = class(TWorker2)
  strict private
    FOutlook,
    FOLMAPINameSpace,
    FOLMAPIFolders: OLEVariant;

    procedure RespondEnqueueAndNotifyMainComm(AMsgId: word;
      const AValue: TOmniValue; const AWinMsg: integer); overload;
    procedure RespondEnqueueAndNotifyMainComm(AMsg: TOmniMessage; const AWinMsg: integer); overload;

    procedure InitMAPI;
    function GetAllOLPublicFolderList: TStringList;
    procedure GetOLFolderList(AFolderKind: integer); overload;
    procedure GetOLFolderList(AMAPIFolder: OLEVariant; AStrList: TStringList); overload;
    procedure GetOLFolderList2TV(tvFolders: TTreeView);

    function GetSelectedMailItemsFromExplorer: RawUTF8; //Json Array 형식으로 반환 함
    procedure ShowMailContents(AEntryId, AStoreId: string);
  protected
    procedure Execute; override;
    procedure ProcessCommandProc(AMsg: TOmniMessage); override;
    procedure ProcessRespondData(AMsg: TOmniMessage);

    procedure ProcessGetFolderList(AMsg: TOmniMessage);
    procedure ProcessGetSelectedMailItemFromExplorer(AMsg: TOmniMessage);
    procedure ProcessShowMailContents(AMsg: TOmniMessage);
  public
    constructor Create(commandQueue, responseQueue, sendQueue: TOmniMessageQueue; AFormHandle: THandle);
    destructor Destroy(); override;

    procedure InitVar();
    procedure Log2MainComm(const AMsg: string);
    procedure CustomCreate; override; //Contructor 보다 먼저 실행 됨
  end;

implementation

{ TOLControlWorker }

constructor TOLControlWorker.Create(commandQueue, responseQueue,
  sendQueue: TOmniMessageQueue; AFormHandle: THandle);
begin
  inherited Create(commandQueue, responseQueue, sendQueue);
  FormHandle := AFormHandle;
  FOutlook := null;

//  InitVar();
end;

procedure TOLControlWorker.CustomCreate;
begin
end;

destructor TOLControlWorker.Destroy;
begin
  if not VarIsNull(FOutlook) then
  begin
    FOutlook := null;
  end;

  inherited;
end;

procedure TOLControlWorker.Execute;
var
  handles: array [0..1] of THandle;
  msg    : TOmniMessage;
begin
  CoInitialize(nil);
  try
    handles[0] := StopEvent.Handle;
    handles[1] := CommandQueue.GetNewMessageEvent;

    while WaitForMultipleObjects(2, @handles, false, INFINITE) = (WAIT_OBJECT_0 + 1) do
    begin
      if terminated then
        break;

      while CommandQueue.TryDequeue(msg) do
      begin
        ProcessCommandProc(msg);
      end;//while
    end;//while
  finally
    CoUninitialize;
  end;
end;

function TOLControlWorker.GetAllOLPublicFolderList: TStringList;
var
  i: integer;
  LMAPIFolder: OLEVariant;//MAPIFolder;
begin
  Result := TStringList.Create;

  for i := 1 to FOLMAPINameSpace.Folders.Count do
  begin
    LMAPIFolder := FOLMAPINameSpace.Folders.Item[i];
    GetOLFolderList(LMAPIFolder, Result);
  end;
end;

procedure TOLControlWorker.GetOLFolderList(AMAPIFolder: OLEVariant; AStrList: TStringList);
var
  i: Integer;
  LMAPISubFolder: OLEVariant;
begin
  if AMAPIFolder.Folders.Count = 0 then
    AStrList.Add(AMAPIFolder.FullFolderPath + '=' + AMAPIFolder.StoreID)
  else
  begin
    for i := 1 to AMAPIFolder.Folders.Count do
    begin
      LMAPISubFolder := AMAPIFolder.Folders.Item[i];
      GetOLFolderList(LMAPISubFolder, AStrList);
    end;
  end;
end;

procedure TOLControlWorker.GetOLFolderList2TV(tvFolders: TTreeView);
var
  node: TTreeNode;

  procedure _LoadFolder(AParentNode: TTreeNode; AFolder: OLEVariant);
  var i: integer;
  begin
    for i := 1 to AFolder.Count do
    begin
      node := tvFolders.Items.AddChild(AParentNode, AFolder.Item[i].Name);

      _LoadFolder(node, AFolder.Item[i].Folders);
    end;
  end;
begin
  _LoadFolder(nil, FOLMAPINameSpace.Folders);
end;

function TOLControlWorker.GetSelectedMailItemsFromExplorer: RawUTF8;
var
  LExplorer,//: _Explorer;
  LSelection,//: Selection;
  LMailItem,//: _MailItem;
  LAddressEntry, //AddressEntry
  LRecipients, //Recipients
  LRecipient, //Recipient
  LFolder //Folder
  : OLEVariant;
  i,j: integer;
  LDynArr: TDynArray;
  LDynUtf8: TRawUTF8DynArray;
  LVar: variant;
  LUtf8: RawUTF8;
  LStr: RawUTF8;
begin
  TDocVariant.New(LVar);
  LDynArr.Init(TypeInfo(TRawUTF8DynArray), LDynUtf8);

  LExplorer := FOutlook.ActiveExplorer;
  LSelection := LExplorer.Selection;

  for i := 1 to LSelection.Count do
  begin
    LMailItem := LSelection.Item(i);
    LVar.LocalEntryId := LMailItem.EntryID;
    LVar.Subject := LMailItem.Subject;
    LVar.SenderEmailAddress := LMailItem.SenderEmailAddress;
    LVar.SenderName := LMailItem.SenderName;
    LVar.CC := LMailItem.CC;
    LVar.BCC := LMailItem.BCC;
    LVar.HTMLBody := LMailItem.HTMLBody;
    LVar.ReceivedTime := LMailItem.ReceivedTime;

    LFolder := LMailItem.Parent;
    LVar.SavedOLFolderPath := LFolder.FullFolderPath;
    LVar.LocalStoreID := LFolder.StoreID;

//    LAddressEntry := LMailItem.AddressEntry;
    LRecipients := LMailItem.Recipients;

    LStr := '';

    for j := 1 to LRecipients.Count do
    begin
      LRecipient := LRecipients.Item(j);
      LStr := LStr + LRecipient.Address + ';';
    end;

    LVar.Recipients := LStr;

    LUtf8 := _JSON(LVar);

    LDynArr.Add(LUtf8);
  end;

  Result := LDynArr.SaveToJson;

end;

procedure TOLControlWorker.GetOLFolderList(AFolderKind: integer);
begin
  FOLMAPIFolders := FOLMAPINameSpace.GetDefaultFolder(AFolderKind); //olFolderInbox
end;

procedure TOLControlWorker.InitMAPI;
var
  i: integer;
  LMAPIFolder: OLEVariant;//MAPIFolder;
  LFolderList: TStringList;
begin
  FOLMAPINameSpace := FOutlook.GetNameSpace('MAPI');
  Log2MainComm('GetNameSpace(''MAPI'')');
  FOLMAPINameSpace.Logon('', '', False, True);
  Log2MainComm('FOLMAPINameSpace.Logon('', '', False, True)');
end;

procedure TOLControlWorker.InitVar;
begin
  if VarIsNull(FOutlook) then //UnAssigned
  begin
    try
      FOutlook := GetActiveOleObject('outlook.application');
      Log2MainComm('OutLook Activated!');
    except
      try
        FOutlook := CreateOleObject('outlook.application');
        Log2MainComm('OutLook Created!');
      except
        // Unable to access or start OUTLOOK
        Log2MainComm(
          'Unable to start or access Outlook.  Possibilities include: permission problems, server down, or VPN not enabled.  Exiting...');
        exit;
      end;
    end;
  end;

  InitMAPI();
end;

procedure TOLControlWorker.Log2MainComm(const AMsg: string);
var
  LValue: TOmniValue;
  LOLRespondRec: TOLRespondRec;
  LOmniMsg: TOmniMessage;
begin
  LOLRespondRec.FID := Ord(olrkLog);
  LOLRespondRec.FMsg := AMsg;
  LValue := TOmniValue.FromRecord(LOLRespondRec);
  LOmniMsg := TOmniMessage.Create(Ord(olrkLog), LValue);

  RespondEnqueueAndNotifyMainComm(LOmniMsg, MSG_RESULT);
end;

procedure TOLControlWorker.ProcessCommandProc(AMsg: TOmniMessage);
begin
  case TOLCommandKind(AMsg.MsgID) of
    olckInitVar: InitVar();
    olckAddAppointment: ;
    olckGetFolderList: begin
      ProcessGetFolderList(AMsg);
    end;
    olckMoveMail2Folder: ;
    olckGetSelectedMailItemFromExplorer: begin
      ProcessGetSelectedMailItemFromExplorer(AMsg);
    end;
    olckShowMailContents: begin
      ProcessShowMailContents(AMsg);
    end;
  end;
end;

procedure TOLControlWorker.ProcessGetFolderList(AMsg: TOmniMessage);
var
  LValue: TOmniValue;
  LFolderList: TStringList;
  LOLRespondRec: TOLRespondRec;
  LOmniMsg: TOmniMessage;
begin
  LFolderList := GetAllOLPublicFolderList();
  try
    LOLRespondRec.FID := AMsg.MsgID;
    LOLRespondRec.FMsg := LFolderList.Text;
    LValue := TOmniValue.FromRecord(LOLRespondRec);
    LOmniMsg := TOmniMessage.Create(Ord(olrkMAPIFolderList), LValue);
    ProcessRespondData(LOmniMsg);
  finally
    LFolderList.Free;
  end;
end;

procedure TOLControlWorker.ProcessGetSelectedMailItemFromExplorer(
  AMsg: TOmniMessage);
var
  LValue: TOmniValue;
  LMailList: RawUtf8;
  LOLRespondRec: TOLRespondRec;
  LOmniMsg: TOmniMessage;
begin
  LMailList := GetSelectedMailItemsFromExplorer();
  try
    LOLRespondRec.FID := AMsg.MsgID;
    LOLRespondRec.FMsg := Utf8ToString(LMailList);
    LValue := TOmniValue.FromRecord(LOLRespondRec);
    LOmniMsg := TOmniMessage.Create(Ord(olrkSelMail4Explore), LValue);
    ProcessRespondData(LOmniMsg);
  finally
  end;
end;

procedure TOLControlWorker.ProcessRespondData(AMsg: TOmniMessage);
begin
  //MainForm에 값을 전달함
  RespondEnqueueAndNotifyMainComm(AMsg, MSG_RESULT);
end;

procedure TOLControlWorker.ProcessShowMailContents(AMsg: TOmniMessage);
var
  LEntryIdRecord: TEntryIdRecord;
  LOLRespondRec: TOLRespondRec;
  LOmniMsg: TOmniMessage;
//  LStore: OLEVariant;//Store;
begin
  LEntryIdRecord := AMsg.MsgData.ToRecord<TEntryIdRecord>;
//  LStore := FOLMAPINameSpace.GetStoreFromID(LEntryIdRecord.FStoreId);
  ShowMailContents(LEntryIdRecord.FEntryId, LEntryIdRecord.FStoreId);

//  LOLRespondRec.FID := AMsg.MsgID;
//  LOLRespondRec.FMsg := Utf8ToString(LMailList);
//  LValue := TOmniValue.FromRecord(LOLRespondRec);
//  LOmniMsg := TOmniMessage.Create(Ord(olrkSelMail4Explore), LValue);
//  ProcessRespondData(LOmniMsg);
end;

procedure TOLControlWorker.RespondEnqueueAndNotifyMainComm(AMsg: TOmniMessage;
  const AWinMsg: integer);
begin
  if ResponseQueue.Enqueue(AMsg) then
    SendMessage(FormHandle, AWinMsg, AWinMsg, 0)
  else
    raise System.SysUtils.Exception.Create('Response queue is full!');
end;

procedure TOLControlWorker.ShowMailContents(AEntryId, AStoreId: string);
var
  LMailItem: OLEVariant;//MailItem;
begin
  LMailItem := FOLMAPINameSpace.GetItemFromID(AEntryId, AStoreId);// as MailItem;

  if not VarIsNull(LMailItem) then
  begin
    LMailItem.Display(False);
  end;
end;

procedure TOLControlWorker.RespondEnqueueAndNotifyMainComm(AMsgId: word;
  const AValue: TOmniValue; const AWinMsg: integer);
begin
  if ResponseQueue.Enqueue(TOmniMessage.Create(AMsgId, AValue)) then
    SendMessage(FormHandle, AWinMsg, AMsgId, 0)
  else
    raise System.SysUtils.Exception.Create('Response queue is full!');
end;

end.
