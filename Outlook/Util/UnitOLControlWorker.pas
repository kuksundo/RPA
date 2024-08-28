unit UnitOLControlWorker;

interface

uses Windows, Winapi.Messages, System.SysUtils, System.SyncObjs, System.Classes,
  Variants, System.Win.ComObj, Vcl.ComCtrls, ActiveX, DateUtils,
  mormot.core.json, mormot.core.data, mormot.core.base, mormot.core.variants,
  mormot.core.text,
  OtlComm, OtlCommon,
  UnitWorker4OmniMsgQ, Outlook_TLB,
  UnitOutLookDataType;

type
  TOLControlWorker = class(TWorker2)
  strict private
    FOutlook,
    FOLMAPINameSpace,
    FOLMAPIFolders,
    FOLCalendarFolder: OLEVariant;

    procedure RespondEnqueueAndNotifyMainComm(AMsgId: word;
      const AValue: TOmniValue; const AWinMsg: integer); overload;
    procedure RespondEnqueueAndNotifyMainComm(AMsg: TOmniMessage; const AWinMsg: integer); overload;

    procedure InitMAPI;

    //ALevelLimit: Folder depth ����, 1 = Root folder���� 1�ܰ� �Ʒ� Folder ������ ��ȯ��
    function GetAllOLPublicFolderList(ALevelLimit: integer=0; AIsOnlyFolderName: Boolean=False): TStringList;
    procedure GetOLFolderList(AFolderKind: integer); overload;
    procedure GetOLFolderList(AMAPIFolder: OLEVariant; AStrList: TStringList; ALevelLimit: integer; AIsOnlyFolderName: Boolean); overload;
    procedure GetOLFolderList2TV(tvFolders: TTreeView);
    //AFolderPath: ';'�� root folder�� subfolder ���� ���е�
    function CheckIfExistFolder(AFolderPath: string): Boolean;
    //AFolderPath: ';'�� root folder�� subfolder ���� ���е�
    function CreateFolder2Path(AFolderPath: TEntryIdRecord): OLEVariant;//Folder;
    function GetOLMAPIFolderList4Recursive(AMAPIFolder: OLEVariant): OLEVariant;
    function GetOLMAPIFolderByFolderName(AMAPIFolder: OLEVariant; AFolderPath: string): OLEVariant;
    //Folder Full Path�� �̿��Ͽ� Folder ��ü ��ȯ��
    function GetFolderObjectFromPath(AFolderPath: string): OLEVariant;
    //AFolderPath: Root Folder Path + ';' + SubFolder Path
    function GetFolderPathFromRootNSubFolder(AFolderPath: string): string;
    //Folder�� ���� Root ���� Full Path�� ��ȯ��
    function GetFolderFullPathByFolderObj(AMAPIFolder: OLEVariant): string;

    function GetSelectedMailItemsFromExplorer: RawUTF8; //Json Array �������� ��ȯ ��
    procedure MoveMail2Folder(AOriginalEntryId, AOriginalStoreId, AFolderPath: string);
    procedure ShowMailContents(AEntryId, AStoreId: string);

    procedure AddObject2OL(var AOLObjRec: TOLObjectRec);
    procedure ShowObject(AOLAppointRec: TOLObjectRec);

    procedure CreateOLMail(AOLMailRec: TOLMailRec);

    procedure AssignItemFieldByOLObjKind(AObjectItem: OLEVariant; AOLObjRec: TOLObjectRec; AItem: Integer);

  protected
    procedure Execute; override;
    procedure ProcessCommandProc(AMsg: TOmniMessage); override;
    procedure ProcessRespondData(AMsg: TOmniMessage);

    procedure ProcessInitOutlook(AMsg: TOmniMessage);
    procedure ProcessGetFolderList(AMsg: TOmniMessage);
    procedure ProcessGetSelectedMailItemFromExplorer(AMsg: TOmniMessage);
    procedure ProcessShowMailContents(AMsg: TOmniMessage);
    procedure ProcessMoveMail2Folder(AMsg: TOmniMessage);
    procedure ProcessAddObject(AMsg: TOmniMessage);
    procedure ProcessShowObject(AMsg: TOmniMessage);
    procedure ProcessCreateMail(AMsg: TOmniMessage);
    procedure ProcessGotoFolder(AMsg: TOmniMessage);
  public
    constructor Create(commandQueue, responseQueue, sendQueue: TOmniMessageQueue; AFormHandle: THandle);
    destructor Destroy(); override;

    procedure InitVar();
    procedure Log2MainComm(const AMsg: string);
    procedure CustomCreate; override; //Contructor ���� ���� ���� ��
  end;

implementation

uses UnitStringUtil, UnitMiscUtil;

{ TOLControlWorker }

procedure TOLControlWorker.AddObject2OL(var AOLObjRec: TOLObjectRec);
var
  LObjectItem:OLEVariant;
  LItem: LongWord;
begin
  if VarIsNull(FOLCalendarFolder) or VarIsEmpty(FOLCalendarFolder) then
  begin
    FOLCalendarFolder := FOLMAPINameSpace.GetDefaultFolder(olFolderCalendar);
  end;

  LItem := GetOLObjItemFromOLKind(AOLObjRec.OLObjectKind);
  LObjectItem := FOutlook.CreateItem(LItem);

  if not VarIsNull(LObjectItem) then
  begin
    try
//      LObjectItem.MeetingStatus := 1; //olMeeting = 1; set to 0 if there are no recipients/attendees
//      LObjectItem.Subject := 'Outlook Meeting Item';
//      LObjectItem.Body := 'This Microsoft Outlook calendar meeting was created programmatically by Delphi!' + #13#10 + 'Calendar meeting invitations were sent to required and optional attendees.';
//      LObjectItem.Location := 'My office';
//      LObjectItem.AllDayEvent := False;
//      LObjectItem.Start := EncodeDateTime(2022, 8, 7, 10, 0, 0, 0);
//      LObjectItem.End := EncodeDateTime(2022, 8, 7, 10, 50, 0, 0);
//      LObjectItem.Recipients.Add('recipient1@example.com'); //change the recipient email address
//      LObjectItem.Recipients.Add('recipient2@example.com'); //change the recipient email address
//      LObjectItem.RequiredAttendees := 'recipient1@example.com'; //change the recipient email address
//      LObjectItem.OptionalAttendees := 'recipient2@example.com'; //change the recipient email address
      AssignItemFieldByOLObjKind(LObjectItem, AOLObjRec, AOLObjRec.OLObjectKind);

      LObjectItem.Save;
//      LObjectItem.Send;

      AOLObjRec.EntryID := LObjectItem.EntryID;
    finally
      LObjectItem := Unassigned;
    end;
  end;
end;

procedure TOLControlWorker.AssignItemFieldByOLObjKind(AObjectItem: OLEVariant;
  AOLObjRec: TOLObjectRec; AItem: Integer);
begin
  AObjectItem.Subject := AOLObjRec.Subject;
  AObjectItem.Body := AOLObjRec.Body;

  case TOLObjectKind(AItem) of
    olobjAppointment: begin
      AObjectItem.Start := AOLObjRec.Start;
      AObjectItem.End := AOLObjRec.End_;
      AObjectItem.Categories := AOLObjRec.Categories;
    end;
    olobjTask: begin
      AObjectItem.StartDate := AOLObjRec.Start;
      AObjectItem.DueDate := AOLObjRec.End_;
      AObjectItem.Categories := AOLObjRec.Categories;
    end;
    olobjMeeting: begin
    end;
    olobjEvent: begin
    end;
    olobjNote: begin
    end;
    olobjContact: begin

    end;
    olobjVCard: begin
    end;
  end;

end;

function TOLControlWorker.CheckIfExistFolder(AFolderPath: string): Boolean;
var
  LFolderList: TStringList;
  LFolderFullName: string;
begin
  LFolderList := GetAllOLPublicFolderList(-1, True);
  try
    LFolderFullName := GetFolderPathFromRootNSubFolder(AFolderPath);
    Result := LFolderList.IndexOf(LFolderFullName) > -1;
  finally
    LFolderList.Free;
  end;
end;

constructor TOLControlWorker.Create(commandQueue, responseQueue,
  sendQueue: TOmniMessageQueue; AFormHandle: THandle);
begin
  inherited Create(commandQueue, responseQueue, sendQueue);
  FormHandle := AFormHandle;
  FOutlook := null;

//  InitVar();
end;

function TOLControlWorker.CreateFolder2Path(AFolderPath: TEntryIdRecord): OLEVariant;
var
  LFolderList: TStringList;
  LRootFolderName, LSubFolderName, LStr,
  LRootEntryId, LRootStoreId: string;
  LExistFolder: Boolean;
  i: integer;
  LMAPIFolder: OLEVariant;//MAPIFolder;
begin
  Result := null;

  LStr := AFolderPath.FFolderPath4Move;
  LRootFolderName := StrToken(LStr, ';');
  LSubFolderName := StrToken(LStr, ';');

  LSubFolderName := LSubFolderName.Replace('/', '\');

  LMAPIFolder := FOLMAPINameSpace.GetFolderFromID(AFolderPath.FEntryId4MoveRoot, AFolderPath.FStoreId4MoveRoot);

  if not VarIsNull(LMAPIFolder) then
  begin
    while LSubFolderName <> '' do
    begin
      LStr := StrToken(LSubFolderName, '\');
      LExistFolder := False;

      for i := 1 to LMAPIFolder.Folders.Count do
      begin
        if LMAPIFolder.Folders.Item[i].Name = LStr then
        begin
          LMAPIFolder := LMAPIFolder.Folders.Item[i];
          LExistFolder := True;
          Break;
        end;
      end;//for

      //LMAPIFolder Root�� Folder�� ������ ������
      if (not LExistFolder) and (LStr <> '') then
      begin
        LMAPIFolder := LMAPIFolder.Folders.Add(LStr);
        Result := LMAPIFolder;
      end;
    end;//while
  end;
end;

procedure TOLControlWorker.CreateOLMail(AOLMailRec: TOLMailRec);
var
  LMailItem:OLEVariant;
begin
  if VarIsNull(FOLCalendarFolder) or VarIsEmpty(FOLCalendarFolder) then
  begin
    FOLCalendarFolder := FOLMAPINameSpace.GetDefaultFolder(olFolderCalendar);
  end;

  LMailItem := FOutlook.CreateItem(olMailItem);

  if not VarIsNull(LMailItem) then
  begin
    try
      LMailItem.Subject := AOLMailRec.Subject;
      LMailItem.To := AOLMailRec.To_;
      LMailItem.BCC := AOLMailRec.BCC;
      LMailItem.CC := AOLMailRec.CC;
      LMailItem.Body := AOLMailRec.Body;
      LMailItem.HTMLBody := AOLMailRec.HTMLBody;

      LMailItem.Display(False);
    finally
      LMailItem := Unassigned;
    end;
  end;
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

function TOLControlWorker.GetAllOLPublicFolderList(ALevelLimit: integer; AIsOnlyFolderName: Boolean): TStringList;
var
  i: integer;
  LMAPIFolder: OLEVariant;//MAPIFolder;
begin
  Result := TStringList.Create;

  for i := 1 to FOLMAPINameSpace.Folders.Count do
  begin
    LMAPIFolder := FOLMAPINameSpace.Folders.Item[i];
    GetOLFolderList(LMAPIFolder, Result, ALevelLimit, AIsOnlyFolderName);
  end;
end;

function TOLControlWorker.GetFolderFullPathByFolderObj(
  AMAPIFolder: OLEVariant): string;
var
  LFolder: OLEVariant;
begin
  Result := AMAPIFolder.Name;

  LFolder := AMAPIFolder.Parent;

//  while (not VarIsNull(LFolder)) and (LFolder.Class_ <> olNamespace) do
//  while (not VarIsNull(LFolder)) and (not VarIsNull(LFolder.Parent)) do
//  while (not VarIsNull(LFolder)) and (LFolder.Session.CurrentUser.Name <> 'kuksundo') do
  while (not VarIsNull(LFolder)) and (not Supports(LFolder, _NameSpace)) do
  begin
    Result := LFolder.Name + '\' + Result;
    LFolder := LFolder.Parent;
  end;
end;

function TOLControlWorker.GetFolderObjectFromPath(
  AFolderPath: string): OLEVariant;
var
  i: integer;
  LFolderName: string;
//  LMAPIFolder: OLEVariant;//MAPIFolder;
begin
  //Folder Path���� '\\' ����
  StrToken(AFolderPath, '\');
  StrToken(AFolderPath, '\');

  LFolderName := StrToken(AFolderPath, '\');

  for i := 1 to FOLMAPINameSpace.Folders.Count do
  begin
    Result := FOLMAPINameSpace.Folders.Item[i];

    if Result.Name = LFolderName then
      Result := GetOLMAPIFolderByFolderName(Result, AFolderPath);

    if not VarIsNull(Result) then
      Break;
  end;
end;

function TOLControlWorker.GetFolderPathFromRootNSubFolder(
  AFolderPath: string): string;
var
  LRootFolderName, LSubFolderName: string;
begin
  LRootFolderName := StrToken(AFolderPath, ';');
  LRootFolderName := IncludeTrailingPathDelimiter(LRootFolderName);

  LSubFolderName := StrToken(AFolderPath, ';');
  TrimLeftChar(LSubFolderName, '\');

  LSubFolderName := LSubFolderName.Replace('/', '\');

  Result := LRootFolderName + LSubFolderName;
end;

procedure TOLControlWorker.GetOLFolderList(AMAPIFolder: OLEVariant;
  AStrList: TStringList; ALevelLimit: integer; AIsOnlyFolderName: Boolean);
var
  i, LvlLimit: Integer;
  LMAPISubFolder: OLEVariant;
  LStr: string;
begin
  //���� ������ ���� ������ ��ȯ��
//  if AMAPIFolder.Folders.Count = 0 then
  //ALevelLimit���� ������ Depth������ ��ȯ��
  if ALevelLimit = 0 then
  begin
    if AIsOnlyFolderName then
      LStr := AMAPIFolder.FullFolderPath
    else
      LStr := AMAPIFolder.FullFolderPath + '=' + AMAPIFolder.EntryID + ';' + AMAPIFolder.StoreID;

    AStrList.Add(LStr);
  end
  else
  begin
    if AIsOnlyFolderName then
      LStr := AMAPIFolder.FullFolderPath
    else
      LStr := AMAPIFolder.FullFolderPath + '=' + AMAPIFolder.EntryID + ';' + AMAPIFolder.StoreID;

    AStrList.Add(LStr);

    Dec(ALevelLimit);

    for i := 1 to AMAPIFolder.Folders.Count do
    begin
      LMAPISubFolder := AMAPIFolder.Folders.Item[i];
      GetOLFolderList(LMAPISubFolder, AStrList, ALevelLimit, AIsOnlyFolderName);
    end;//for
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

function TOLControlWorker.GetOLMAPIFolderByFolderName(AMAPIFolder: OLEVariant;
  AFolderPath: string): OLEVariant;
var
  i: integer;
  LFolderName: string;
  LMAPIFolder: OLEVariant;
begin
  Result := null;

  if AFolderPath = '' then
    Result := AMAPIFolder
  else
  begin
    LFolderName := StrToken(AFolderPath, '\');

    if LFolderName <> '' then
    begin
      for i := 1 to AMAPIFolder.Folders.Count do
      begin
        LMAPIFolder := AMAPIFolder.Folders.Item[i];

        if LMAPIFolder.Name = LFolderName then
        begin
          LMAPIFolder := GetOLMAPIFolderByFolderName(LMAPIFolder, AFolderPath);
          Result := LMAPIFolder;
          Break;
        end;
      end;
    end;
  end;
end;

function TOLControlWorker.GetOLMAPIFolderList4Recursive(
  AMAPIFolder: OLEVariant): OLEVariant;
var
  i: Integer;
  LMAPISubFolder: OLEVariant;
begin
  if AMAPIFolder.Folders.Count = 0 then
    Result := AMAPIFolder
  else
  begin
    for i := 1 to AMAPIFolder.Folders.Count do
    begin
      LMAPISubFolder := AMAPIFolder.Folders.Item[i];
      LMAPISubFolder := GetOLMAPIFolderList4Recursive(LMAPISubFolder);
    end;
  end;
end;

function TOLControlWorker.GetSelectedMailItemsFromExplorer: RawUTF8;
var
  LExplorer,//: _Explorer;
  LSelection,//: Selection;
  LMailItem,//: _MailItem;
  LAddressEntry, //AddressEntry
  LRecipients, //Recipients
  LRecipient, //Recipient
  LAttachments, //Attachments
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

    //Item Name�� grid_Mail Column Name�� ��ġ�ؾ� ��
    LVar.LocalEntryId := LMailItem.EntryID;
    LVar.Subject := LMailItem.Subject;
    LVar.SenderEmail := LMailItem.SenderEmailAddress;
    LVar.SenderName := LMailItem.SenderName;
    LVar.CC := LMailItem.CC;
    LVar.BCC := LMailItem.BCC;
    LVar.HTMLBody := LMailItem.HTMLBody;
    LVar.RecvDate := LMailItem.ReceivedTime;//VarFromDateTime()

    LFolder := LMailItem.Parent;
    LVar.SavedOLFolderPath := LFolder.FullFolderPath;
    LVar.LocalStoreId := LFolder.StoreID;
    LVar.FolderEntryId := LFolder.EntryID;
    LRecipients := LMailItem.Recipients;

    LAttachments := LMailItem.Attachments;
    LVar.AttachCount := LAttachments.Count;

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
//      Log2MainComm('OutLook Activated!');
    except
      try
        FOutlook := CreateOleObject('outlook.application');
//        Log2MainComm('OutLook Created!');
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

procedure TOLControlWorker.MoveMail2Folder(AOriginalEntryId, AOriginalStoreId,
  AFolderPath: string);
begin

end;

procedure TOLControlWorker.ProcessAddObject(AMsg: TOmniMessage);
var
  LOLAppointRec: TOLObjectRec;
  LOLRespondRec: TOLRespondRec;
  LOmniMsg: TOmniMessage;
  LValue: TOmniValue;
begin
  LOLAppointRec := AMsg.MsgData.ToRecord<TOLObjectRec>;
  FormHandle := LOLAppointRec.FSenderHandle;

  //outlook�� ���: EntryId�� LOLAppointRec.EntryId�� ä����
  AddObject2OL(LOLAppointRec);

  LOLRespondRec.FID := AMsg.MsgID;
  LOLRespondRec.FMsg := RecordSaveJson(LOLAppointRec, TypeInfo(TOLObjectRec));
  LOLRespondRec.FSenderHandle := FormHandle;

  LValue := TOmniValue.FromRecord(LOLRespondRec);
  LOmniMsg := TOmniMessage.Create(Ord(olrkAddObject), LValue);

  ProcessRespondData(LOmniMsg);
end;

procedure TOLControlWorker.ProcessCommandProc(AMsg: TOmniMessage);
begin
  case TOLCommandKind(AMsg.MsgID) of
    olckInitVar: begin
      ProcessInitOutlook(AMsg);
    end;
    olckAddObject: ProcessAddObject(AMsg);
    olckGetFolderList: begin
      ProcessGetFolderList(AMsg);
    end;
    olckMoveMail2Folder: begin
      ProcessMoveMail2Folder(AMsg);
    end;
    olckGetSelectedMailItemFromExplorer: begin
      ProcessGetSelectedMailItemFromExplorer(AMsg);
    end;
    olckShowMailContents: begin
      ProcessShowMailContents(AMsg);
    end;
    olckShowObject: begin
      ProcessShowObject(AMsg);
    end;
    olckCreateMail: begin
      ProcessCreateMail(AMsg);
    end;
    olcGotoFolder: begin
      ProcessGotoFolder(AMsg);
    end;
  end;
end;

procedure TOLControlWorker.ProcessCreateMail(AMsg: TOmniMessage);
var
  LOLMailRec: TOLMailRec;
  LOLRespondRec: TOLRespondRec;
  LOmniMsg: TOmniMessage;
begin
  LOLMailRec := AMsg.MsgData.ToRecord<TOLMailRec>;
  FormHandle := LOLMailRec.FSenderHandle;

  CreateOLMail(LOLMailRec);
end;

procedure TOLControlWorker.ProcessGetFolderList(AMsg: TOmniMessage);
var
  LValue: TOmniValue;
  LFolderList: TStringList;
  LOLRespondRec: TOLRespondRec;
  LOmniMsg: TOmniMessage;
begin
  LFolderList := GetAllOLPublicFolderList(2);
  try
    FormHandle := AMsg.MsgData.AsInteger;

    LOLRespondRec.FID := AMsg.MsgID;
    LOLRespondRec.FMsg := LFolderList.Text;
    LOLRespondRec.FSenderHandle := FormHandle;

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
  //Outlook���� Select�� MailList�� Array Json���·� ��ȯ��
  //grid_Mail Column Name�� ������ Name��
  LMailList := GetSelectedMailItemsFromExplorer();
  try
    //FSenderHandle�� ����
    LOLRespondRec := AMsg.MsgData.ToRecord<TOLRespondRec>;
    FormHandle := LOLRespondRec.FSenderHandle;
    LOLRespondRec.FID := AMsg.MsgID;
    LOLRespondRec.FMsg := Utf8ToString(LMailList);
    LValue := TOmniValue.FromRecord(LOLRespondRec);
    LOmniMsg := TOmniMessage.Create(Ord(olrkSelMail4Explore), LValue);
    ProcessRespondData(LOmniMsg);
  finally
  end;
end;

procedure TOLControlWorker.ProcessGotoFolder(AMsg: TOmniMessage);
var
  LFolder   //MAPIFolder
  : OLEVariant;
  LEntryIdRecord: TEntryIdRecord;
  LFolderPath: string;
begin
  LEntryIdRecord := AMsg.MsgData.ToRecord<TEntryIdRecord>;

  //FFolderPath4Move : '\\great.park@hd.com;SHI8151\098'
  if CheckIfExistFolder(LEntryIdRecord.FFolderPath4Move) then
  begin
    //�����ݷ��� ���ְ� "\" �߰��Ͽ� �ϳ��� ��ħ (\\great.park@hd.com\SHI8151\098)
    LFolderPath := GetFolderPathFromRootNSubFolder(LEntryIdRecord.FFolderPath4Move);
    LFolder := GetFolderObjectFromPath(LFolderPath);

    FOutlook.ActiveExplorer.SelectFolder(LFolder);
  end
end;

procedure TOLControlWorker.ProcessInitOutlook(AMsg: TOmniMessage);
var
  LOmniMsg: TOmniMessage;
  LValue: TOmniValue;
  LOLRespondRec: TOLRespondRec;
begin
  FormHandle := AMsg.MsgData.AsInteger;
  InitVar();

//  LOLRespondRec.FSenderHandle :=
  LValue := TOmniValue.FromRecord(LOLRespondRec);
//  LOmniMsg := TOmniMessage.Create(Ord(olrkMAPIFolderList), LValue);

  ProcessRespondData(LOmniMsg);
end;

procedure TOLControlWorker.ProcessMoveMail2Folder(AMsg: TOmniMessage);
var
  LDict: IDocDict;
  LMailItem,//MailItem
  LFolder   //Folder
  : OLEVariant;
  LEntryIdRecord: TEntryIdRecord;

  LOLRespondRec: TOLRespondRec;
  LOmniMsg: TOmniMessage;
  LValue: TOmniValue;
  LFolderPath: string;
begin
  //LDoc : {grid_Mail Column Name, vaule} �� Json ������
//  LDict.Json := AMsg.MsgData.AsString;

  LEntryIdRecord := AMsg.MsgData.ToRecord<TEntryIdRecord>;

  LMailItem := FOLMAPINameSpace.GetItemFromID(LEntryIdRecord.FEntryId, LEntryIdRecord.FStoreId);
  //FFolderPath4Move : '\\great.park@hd.com;SHI8151\098'
  if CheckIfExistFolder(LEntryIdRecord.FFolderPath4Move) then
  begin
    //�����ݷ��� ���ְ� "\" �߰��Ͽ� �ϳ��� ��ħ (\\great.park@hd.com\SHI8151\098)
    LFolderPath := GetFolderPathFromRootNSubFolder(LEntryIdRecord.FFolderPath4Move);
    LFolder := GetFolderObjectFromPath(LFolderPath);
  end
  else
    LFolder := CreateFolder2Path(LEntryIdRecord);

  if (not VarIsNull(LMailItem)) and (not VarIsNull(LFolder)) then
  begin
    LDict := DocDict('{}');
    LDict.U['OldEntryId'] := LMailItem.EntryId;

    LMailItem := LMailItem.Move(LFolder);

    LFolder := LMailItem.Parent;

    LDict.U['NewEntryId'] := LMailItem.EntryId;
    LDict.U['NewStoreId'] := LFolder.StoreId;
    LDict.U['NewEntryId4Folder'] := LFolder.EntryId;
    LDict.U['SavedOLFolderPath'] := GetFolderFullPathByFolderObj(LFolder);

    LOLRespondRec.FID := Ord(olrkMoveMail2Folder);
    //�̵��� Mail�� EntryId�� StoreId�� ������
    LOLRespondRec.FMsg := LDict.ToJson(jsonUnquotedPropNameCompact);
    LOLRespondRec.FSenderHandle := LEntryIdRecord.FSenderHandle;
    FormHandle := LEntryIdRecord.FSenderHandle;
//    LOLRespondRec.FMsg := '{"NewEntryId"=' + LMailItem.EntryId + ',"NewStoreId"=' + LMailItem.StoreId + '}';
    LValue := TOmniValue.FromRecord(LOLRespondRec);
    LOmniMsg := TOmniMessage.Create(Ord(olrkMoveMail2Folder), LValue);
    ProcessRespondData(LOmniMsg);
  end;

//  LDict['LocalEntryId'];
//  LDict['LocalStoreId'];
end;

procedure TOLControlWorker.ProcessRespondData(AMsg: TOmniMessage);
begin
  //MainForm�� ���� ������
  RespondEnqueueAndNotifyMainComm(AMsg, MSG_RESULT);
end;

procedure TOLControlWorker.ProcessShowObject(AMsg: TOmniMessage);
var
  LOLAppointRec: TOLObjectRec;
//  LOLRespondRec: TOLRespondRec;
//  LOmniMsg: TOmniMessage;
begin
  LOLAppointRec := AMsg.MsgData.ToRecord<TOLObjectRec>;
//  FormHandle := LOLAppointRec.FSenderHandle;

  ShowObject(LOLAppointRec);
end;

procedure TOLControlWorker.ProcessShowMailContents(AMsg: TOmniMessage);
var
  LEntryIdRecord: TEntryIdRecord;
  LOLRespondRec: TOLRespondRec;
  LOmniMsg: TOmniMessage;
//  LStore: OLEVariant;//Store;
begin
  LEntryIdRecord := AMsg.MsgData.ToRecord<TEntryIdRecord>;
  FormHandle := LEntryIdRecord.FSenderHandle;

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

procedure TOLControlWorker.ShowObject(AOLAppointRec: TOLObjectRec);
var
  LAppointItem: OLEVariant;//ObjectItem;
begin
  LAppointItem := FOutlook.Session.GetItemFromID(AOLAppointRec.EntryID);// as ObjectItem;

  if not VarIsNull(LAppointItem) then
  begin
    LAppointItem.Display(False);
  end;
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
