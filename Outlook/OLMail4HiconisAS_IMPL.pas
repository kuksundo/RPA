unit OLMail4HiconisAS_IMPL;

interface

uses
  SysUtils, ComObj, ComServ, ActiveX, Variants, Outlook2010, Office2010, adxAddIn, OLMail4InqManage_TLB,
  System.Types, System.StrUtils, System.Classes, Dialogs, System.Generics.Collections,
  Winapi.Windows, System.SyncObjs, adxHostAppEvents, adxolFormsManager, Messages,
  AdvAlertWindow, Vcl.ExtCtrls, IdMessage, UnitSynLog2,
  mormot.core.base, mormot.rest.server, mormot.orm.core, mormot.rest.http.server,
  mormot.soa.server, mormot.core.datetime, mormot.rest.memserver, mormot.soa.core,
  mormot.core.interfaces, mormot.core.buffers, mormot.core.unicode, mormot.core.os,
  mormot.core.data, mormot.core.variants, mormot.core.json,
  StompClient, StompTypes,
  OtlCommon, OtlComm, OtlTaskControl, OtlContainerObserver, otlTask,
  Cromis.Comm.Custom, Cromis.Comm.IPC, Cromis.Threading, Cromis.AnyValue,
  OLMailWSCallbackInterface2, UnitCommonWSInterface2, UnitClientInfoClass2, CommonData2,
  UnitOLDataType;

const
  OLUSERID = 'A379042';
  OLMYMAILADDR = 'great.park@hd.com';
//  OLMYMAILADDR2 = 'junghyunpark@hyundai-gs.com';
  OLMYRECVFOLDERPATH = '\\great.park@hd.com\받은 편지함';
  WM_RUN_TASK = WM_USER + 1;
//  AUTOT_FORWARD_MAIL_ACCOUNT = 'jhpark@hyundai-gs.com';

type
  TCoOLMail4InqManage = class(TadxAddin, ICoOLMail4InqManage)
  end;

  TWorker4STOMP = class(TThread)
  private
    FWorker4STOMPQueue: TOmniMessageQueue;
    FWorker4STOMPStopEvent    : TEvent;
    FAutoForwardFolderPathDic: TDictionary<string, MAPIFolder>;
  protected
    procedure Execute; override;
    procedure SendMailToMsgFileThread(AEntryIDList: WideString;
      AStompClient: TStompClient; AHostAddr: string);
    procedure GetFolderPath2Dic(AFolderPath: string);
  public
    constructor Create(sendQueue: TOmniMessageQueue);
    destructor Destroy; override;
    procedure Stop;
  end;

  TWorker4OLMsg = class(TThread)
  private
    FOLMsgQueue4Worker: TOmniMessageQueue;
    FOLMsg2IPCMQ4Worker: TOmniMessageQueue;
    FStopEvent    : TEvent;
    FInboxStoreIdList: TStringList;
    FNameSpace_Worker4OLMsg: _NameSpace;
  protected
    procedure Execute; override;
  public
    constructor Create(sendQueue, IPCQueue: TOmniMessageQueue;
      AInboxStoreIdList: TStringList; ANameSpace: _NameSpace);
    destructor Destroy; override;
    procedure Stop;
  end;

//  TOLMsgFileRecords = array of TOLMsgFileRecord;
//  POLMsgFileRecords = ^TOLMsgFileRecords;

  TServiceOL4WS = class(TInterfacedObject, IOLMailService)
  private
  protected
    fConnected: array of IOLMailCallback;
    FClientInfoList: TStringList;
  public
    constructor Create;
    destructor Destroy; override;

    procedure Join(const pseudo: string; const callback: IOLMailCallback);
    procedure CallbackReleased(const callback: IInvokable; const interfaceName: RawUTF8);
    function ServerExecute(const Acommand: string): RawUTF8;
    function GetOLEmailInfo(ACommand: string): RawUTF8;
    function GetOLEmailAccountInfo: RawUTF8;
  end;

  TAddInModule = class(TadxCOMAddInModule)
    adxContextMenu1: TadxContextMenu;
    adxContextMenu2: TadxContextMenu;
    adxOutlookAppEvents1: TadxOutlookAppEvents;
    AdvAlertWindow1: TAdvAlertWindow;
    Timer1: TTimer;
    procedure adxCOMAddInModuleCreate(Sender: TObject);
    procedure adxCOMAddInModuleDestroy(Sender: TObject);
    procedure adxCOMAddInModuleAddInInitialize(Sender: TObject);
    procedure adxContextMenu1BeforeAddControls(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure adxContextMenu1Controls0Controls0Click(Sender: TObject);
    procedure adxContextMenu1Controls0Controls1Click(Sender: TObject);
    procedure adxContextMenu1Controls0Controls2Click(Sender: TObject);
    procedure adxContextMenu1Controls0Controls3Click(Sender: TObject);
    procedure adxContextMenu1Controls0Controls4Click(Sender: TObject);
    procedure adxContextMenu1Controls1Controls0Click(Sender: TObject);
    procedure adxContextMenu1Controls1Controls2Click(Sender: TObject);
    procedure adxContextMenu1Controls1Controls3Click(Sender: TObject);
    procedure adxContextMenu2Controls0Controls0Click(Sender: TObject);
    procedure adxContextMenu2Controls0Controls2Click(Sender: TObject);
    procedure adxContextMenu2Controls0Controls3Click(Sender: TObject);
    procedure adxOutlookAppEvents1NewMailEx(ASender: TObject;
      const EntryIDCollection: WideString);
    procedure adxContextMenu2Controls0Controls4Click(Sender: TObject);
    procedure adxContextMenu2Controls0Controls5Click(Sender: TObject);
    procedure adxCOMAddInModuleAddInFinalize(Sender: TObject);
  private
    FWorker4STOMP: TWorker4STOMP;
    FWorker4OLMsg: TWorker4OLMsg;
    FOLMsgQueue,
    FOLMsg2IPCMQ,
    FOLMsg2STOMPMQ: TOmniMessageQueue;
    FEmailDisplayMQ: TOmniMessageQueue;
    FEmailDisplayStopEvent    : TEvent;
    FInboxStoreIdList,
//    FAutoForwardFolderPathList,
    FProductTypeStoreIdList: TStringList;

    FIPCClientList: TCommContextList;
    FNameSpace: _NameSpace;
    FEntryIDList: TStringList;
    FAutoSend4OLMsg2IPCMQ,
    FInitCompleted: Boolean;
    FCommModes: TCommModes;

    //Websocket-b
    procedure CreateHttpServer4WS(APort, ATransmissionKey: string;
      aClient: TInterfacedClass; const aInterfaces: array of TGUID);
    procedure DestroyHttpServer;
    function SessionCreate(Sender: TSQLRestServer; Session: TAuthSession;
                  Ctxt: TSQLRestServerURIContext): boolean;
    function SessionClosed(Sender: TSQLRestServer; Session: TAuthSession;
                  Ctxt: TSQLRestServerURIContext): boolean;
    //Websocket-e
  protected
    //Websocket-b
    FModel: TSQLModel;
    FHTTPServer: TSQLHttpServer;
    FRestServer: TSQLRestServer;
    FServiceFactoryServer: TServiceFactoryServer;

    FIpAddr: string;
    FURL: string; //Server에서 Client에 Config Change Notify 하기 위한 Call Back URL
    FIsServerActive,
    FIsPJHPC: Boolean;
    FPortName,
    FUserEmail,
    FUserName: string;
    //Websocket-e

    procedure ShowStoreIdFromSelected;
    procedure GetInboxList;
    procedure GetAutoForwardFolderAccounList;
    procedure DoItemAdd(ASender: TObject; const Item: IDispatch);
    procedure SendEntryID2IPCFromList;
    procedure AdjustDisplayMenu;

    procedure DeleteMsgFilesFromTempFolder;

    procedure InitNetwork;
  public
    //Aflag = 'A'; //쪽지
    //      = 'B'; //SMS
    //AUser = 사번
    procedure Send_Message(AHead, ATitle, AContent, ASendUser, ARecvUser, AFlag: string);
    procedure SendMailToMsgFile(AEntryIDList: WideString);
    procedure SendMailToMsgFile_Async(AEntryIDList: WideString);
    procedure MoveMail2WorkingFolder(AEntryIDList: WideString);
    procedure AssignMailItem2Rec(AMail: MailItem; out ARec: TOLMsgFileRecord);
    procedure MoveMail2Folder(var AEntryIdRecord: TEntryIdRecord);
    function GetFolderFromPath(APath:string): MAPIFOLDER;
    function IsExistFolder(var AFolder: MAPIFolder; AFolderName: string): Boolean;

    function AssignOLMailItemToIdMessage(AOLMailItem: MailItem;
      out AIdMsg: TIdMessage): boolean;
    procedure AssignOLRecipientToIdMsg(AOLMailItem: MailItem;
      out AIdMsg: TIdMessage);
    procedure AssignOLAttachmentToIdMsg(AOLMailItem: MailItem;
      out AIdMsg: TIdMessage);
    procedure ShowMailContents(AEntryId, AStoreId: string);
    procedure ReplyMail(AEntryIdRecord: TEntryIdRecord);
    procedure CreateMail(AEntryIdRecord: TEntryIdRecord);
    procedure ForwardMail(AEntryIdRecord: TEntryIdRecord);
    procedure ViewMailFromMsgFile(AFileName: string);

    procedure CreateAppointment(ATodoItem: variant);

    procedure AsyncEmailDisplay;
    procedure AsyncSendEntryId2IPC(AIsDrag: Boolean=True);
    function GetEmail2StrList: TStringList;
    function GetResponse4MoveFolder2StrList(AEntryIdRecord: TEntryIdRecord): TStringList;

    function ProcessCommandFromClient(ACommand: string): RawUTF8;
    function ServerExecuteFromClient(ACommand: string): RawUTF8;
    function GetOLEmailAccountInfo: RawUTF8;
  end;

    procedure Log4OL(Amsg: string; AIsSaveLog: Boolean = False;
      AMsgLevel: TSynLogInfo = sllInfo);
var
  g_WorkingFolder: MAPIFolder;
  g_PrevStoreId: string;
  MyAddInModule : TAddInModule;

implementation

uses IdEMailAddress, IdAttachmentFile, OtlParallel, Clipbrd, ShellAPI;

{$R *.dfm}

procedure Log4OL(Amsg: string; AIsSaveLog: Boolean;
  AMsgLevel: TSynLogInfo);
begin
  if AIsSaveLog then
  begin
//    Log.Debug(Amsg, '');
//    DoLog(AMsg, True);
  end;
end;

procedure TAddInModule.AdjustDisplayMenu;
begin
  adxContextMenu1.Controls[1].Visible := FIsPJHPC;
end;

procedure TAddInModule.adxCOMAddInModuleAddInFinalize(Sender: TObject);
begin
//  ShowMessage('AddInFinalize');
end;

procedure TAddInModule.adxCOMAddInModuleAddInInitialize(Sender: TObject);
begin
//  ShowMessage('AddInModuleAddInInitialize');
end;

procedure TAddInModule.adxCOMAddInModuleCreate(Sender: TObject);
begin
//  ShowMessage('AddInModuleCreate');
//  InitSynLog;
  Log4OL('adxCOMAddInModuleAddInInitialize', True);
  FOLMsgQueue := TOmniMessageQueue.Create(1000);
  FOLMsg2IPCMQ := TOmniMessageQueue.Create(1000);
  FOLMsg2STOMPMQ := TOmniMessageQueue.Create(1000);
  FEmailDisplayMQ := TOmniMessageQueue.Create(1000);
  FEmailDisplayStopEvent := TEvent.Create;
  FIPCClientList := TCommContextList.Create;
  FInboxStoreIdList := TStringList.Create;
  FProductTypeStoreIdList := TStringList.Create;
  FEntryIDList := TStringList.Create;
//  Log4OL('adxCOMAddInModuleCreate', True);
  FWorker4OLMsg := nil;
  MyAddInModule := Self;
end;

procedure TAddInModule.adxCOMAddInModuleDestroy(Sender: TObject);
var
  i: integer;
begin
//  ShowMessage('AddInModuleDestroy');
  DestroyHttpServer;

  if FAutoSend4OLMsg2IPCMQ then
    SendEntryID2IPCFromList;

  if Assigned(FWorker4OLMsg) then
  begin
    FEmailDisplayStopEvent.SetEvent;
    FWorker4OLMsg.Terminate;
    FWorker4OLMsg.Stop;
    FWorker4STOMP.Terminate;
    FWorker4STOMP.Stop;
    FOLMsgQueue.Free;
    FOLMsg2IPCMQ.Free;
    FOLMsg2STOMPMQ.Free;
    FEmailDisplayMQ.Free;
    FInboxStoreIdList.Free;
    FProductTypeStoreIdList.Free;
    FIPCClientList.Free;
    FEntryIDList.Free;

    FreeAndNil(FEmailDisplayStopEvent);

    DeleteMsgFilesFromTempFolder;
  end;
  Log4OL('adxCOMAddInModuleDestroy', True);
end;

procedure TAddInModule.adxContextMenu1BeforeAddControls(Sender: TObject);
begin
  TadxCommandBarPopup(adxContextMenu1.Controls[1]).Controls[2].Enabled := not FAutoSend4OLMsg2IPCMQ;
  TadxCommandBarPopup(adxContextMenu1.Controls[1]).Controls[3].Enabled := FAutoSend4OLMsg2IPCMQ;
end;

procedure TAddInModule.adxContextMenu1Controls0Controls0Click(Sender: TObject);
begin
  ShowMessage('Add To DPMS To-Do List');
end;

procedure TAddInModule.adxContextMenu1Controls0Controls1Click(Sender: TObject);
begin
//  ShowMessage(IntToStr(FIPCClientList.Count));
end;

procedure TAddInModule.adxContextMenu1Controls0Controls2Click(Sender: TObject);
begin
  ShowMessage(IntToStr(FEntryIDList.Count));
end;

procedure TAddInModule.adxContextMenu1Controls0Controls3Click(Sender: TObject);
var
  LMailItem: MailItem;
  i: integer;
  LOmniValue: TOmniValue;
  LEntryRec: TEntryIdRecord;
  LText: string;
  LFolder: MAPIFolder;
begin
  i := OutlookApp.ActiveExplorer.Selection.Count;

  if i > 1 then
  begin
    ShowMessage('이 기능을 이용하기 위해서는 메일을 1개만 선택 하세요');
    exit;
  end;

  LMailItem := OutlookApp.ActiveExplorer.Selection.Item(i) as MailItem;
  LFolder := OutlookApp.ActiveExplorer.CurrentFolder as MAPIFolder;

  if Assigned(LMailItem) then
  begin
    ShowMessage(LMailItem.EntryID + #13#10 + LFolder.StoreID);
  end;
end;

procedure TAddInModule.adxContextMenu1Controls0Controls4Click(Sender: TObject);
var
  LMailItem: MailItem;
  i: integer;
  LFolder: MAPIFolder;
  LEntryRec: TEntryIdRecord;
  LOmniValue: TOmniValue;
begin
  Parallel.Async(
    procedure (const task: IOmniTask)
    begin
      i := OutlookApp.ActiveExplorer.Selection.Count;

      if i > 1 then
      begin
        ShowMessage('이 기능을 이용하기 위해서는 메일을 1개만 선택 하세요');
        exit;
      end;

      LMailItem := OutlookApp.ActiveExplorer.Selection.Item(i) as MailItem;
      LFolder := OutlookApp.ActiveExplorer.CurrentFolder as MAPIFolder;

      LEntryRec.FEntryId := LMailItem.EntryID;
      LEntryRec.FStoreId := LFolder.StoreID;

//      LEntryRec.FAttachFileName := 'c:\temp\test.msg';
      LOmniValue := TOmniValue.FromRecord<TEntryIdRecord>(LEntryRec);
      FEmailDisplayMQ.Enqueue(TOmniMessage.Create(5, LOmniValue));
    end );
end;

procedure TAddInModule.adxContextMenu1Controls1Controls0Click(Sender: TObject);
var
  LMailItem: MailItem;
  i: integer;
  LOmniValue: TOmniValue;
  LEntryRec: TEntryIdRecord;
  LText: string;
  LFolder: MAPIFolder;
begin
  for i := 1 to OutlookApp.ActiveExplorer.Selection.Count do
  begin
    LMailItem := OutlookApp.ActiveExplorer.Selection.Item(i) as MailItem;
    LFolder := OutlookApp.ActiveExplorer.CurrentFolder as MAPIFolder;

    if Assigned(LMailItem) then
    begin
      LEntryRec.FEntryId := LMailItem.EntryID;
      LEntryRec.FStoreId := LFolder.StoreID;
      LEntryRec.FIgnoreReceiver2pjh := True;
      LEntryRec.FIgnoreEmailMove2WorkFolder := True;
      LOmniValue := TOmniValue.FromRecord<TEntryIdRecord>(LEntryRec);

      if not FOLMsgQueue.Enqueue(TOmniMessage.Create(3, LOmniValue)) then
        Log4OL('Send queue is full!', True)
      else
      begin
        LText := FEntryIDList.Text;
        StringReplace(LText, LMailItem.EntryID, '', [rfReplaceAll]);
        FEntryIDList.Text := LText;
        Log4OL('Send queue is success!', True);
      end;
    end;
  end;//for
end;

procedure TAddInModule.adxContextMenu1Controls1Controls2Click(Sender: TObject);
begin
  FAutoSend4OLMsg2IPCMQ := True;
end;

procedure TAddInModule.adxContextMenu1Controls1Controls3Click(Sender: TObject);
begin
  FAutoSend4OLMsg2IPCMQ := False;
end;

procedure TAddInModule.adxContextMenu2Controls0Controls0Click(Sender: TObject);
var
  LOmniValue: TOmniValue;
  LEntryRec: TEntryIdRecord;
  LFolder: MAPIFolder;
  LText: string;
begin
  LFolder := OutlookApp.ActiveExplorer.CurrentFolder as MAPIFolder;

  LEntryRec.FEntryId := LFolder.FolderPath + ';' + LFolder.Name;
  LEntryRec.FStoreId := LFolder.StoreID;
  LEntryRec.FIgnoreReceiver2pjh := True;
  LEntryRec.FIgnoreEmailMove2WorkFolder := True;
  LOmniValue := TOmniValue.FromRecord<TEntryIdRecord>(LEntryRec);

  if not FOLMsgQueue.Enqueue(TOmniMessage.Create(4, LOmniValue)) then
    Log4OL('Send queue is full!', True)
  else
  begin
    LText := FEntryIDList.Text;
//    StringReplace(LText, LMailItem.EntryID, '', [rfReplaceAll]);
    FEntryIDList.Text := LText;
    Log4OL('Send queue is success!', True);
  end;
end;

procedure TAddInModule.adxContextMenu2Controls0Controls2Click(Sender: TObject);
begin
  ShowStoreIdFromSelected;
end;

procedure TAddInModule.adxContextMenu2Controls0Controls3Click(Sender: TObject);
var
  i: integer;
  LStr: string;
begin
  for i := 1 to FNameSpace.Accounts.Count do
    LStr := LStr + IntToStr(i) + ' : ' + FNameSpace.Accounts.Item(i).SmtpAddress + //FNameSpace.Accounts.Item(i).DisplayName
      ', ' + FNameSpace.Accounts.Item(i).UserName + #13#10;

  ShowMessage(LStr);
end;

procedure TAddInModule.adxContextMenu2Controls0Controls4Click(Sender: TObject);
var
  i,j,k: integer;
  LFolders: _Folders;
begin
  ShowMessage(IntToStr(FNameSpace.Folders.Count));

  for j := 1 to FNameSpace.Folders.Count do
  begin
    LFolders := FNameSpace.Folders.Item(j).Folders;
    for k := 1 to LFolders.Count do
      ShowMessage(LFolders.Item(k).FolderPath + '(' + IntToStr(LFolders.Count) + ')');
  end;
end;

procedure TAddInModule.adxContextMenu2Controls0Controls5Click(Sender: TObject);
begin
  ShowMessage(GetOLEmailAccountInfo);
end;

procedure TAddInModule.adxOutlookAppEvents1NewMailEx(ASender: TObject;
  const EntryIDCollection: WideString);
var
  LOmniValue: TOmniValue;
//  LEntryRec: TEntryIdRecord;
  LRec    : TOLMsgFile4STOMP;
begin
  if not FInitCompleted then
    exit;

  if not FAutoSend4OLMsg2IPCMQ then
  begin
    FEntryIDList.Add(EntryIDCollection);
    exit;
  end;

//  SendMailToMsgFile_Async(EntryIDCollection);
//  SendMailToMsgFile(EntryIDCollection);
  LRec.FHost := MQ_SERVER_IP;
  LRec.FUserId := MQ_USER_ID;
  LRec.FPasswd := MQ_PASSWORD;
  LRec.FMsgFile := EntryIDCollection;
  LOmniValue := TOmniValue.FromRecord<TOLMsgFile4STOMP>(LRec);
//  Log4OL('FOLMsg2STOMPMQ.Enqueue on NewMailEx()!', True);

  if not FOLMsg2STOMPMQ.Enqueue(TOmniMessage.Create(1, LOmniValue)) then
    Log4OL('Send queue is full!', True)
  else
    Log4OL('Send queue is success!', True);
//  if FIPCClientList.Count > 0 then
//  begin
//    LEntryRec.FEntryId := EntryIDCollection;
//    LEntryRec.FStoreId := '';
//    LEntryRec.FIgnoreReceiver2pjh := False;
//    LOmniValue := TOmniValue.FromRecord<TEntryIdRecord>(LEntryRec);
//
//    if not FOLMsgQueue.Enqueue(TOmniMessage.Create(1, LOmniValue)) then
//      Log4OL('Send queue is full!', True)
//    else
//      Log4OL('Send queue is success!', True);
//  end
//  else
//  begin
//    FEntryIDList.Add(EntryIDCollection);
//  end;
end;

procedure TAddInModule.AssignMailItem2Rec(AMail: MailItem;
  out ARec: TOLMsgFileRecord);
var
  i: integer;
begin
  ARec.FSender := AMail.SenderEmailAddress;

  for i := 1 to AMail.Recipients.Count do
    if AMail.Recipients.Item(i).type_ = OlTo then
      ARec.FReceiver := ARec.FReceiver + AMail.Recipients.Item(i).Address + ';'
    else if AMail.Recipients.Item(i).type_ = olCC then
      ARec.FCarbonCopy := ARec.FCarbonCopy + AMail.Recipients.Item(i).Address + ';'
    else if AMail.Recipients.Item(i).type_ = olBCC then
      ARec.FBlindCC := ARec.FBlindCC + AMail.Recipients.Item(i).Address + ';';

  ARec.FSubject := AMail.Subject;
  ARec.FMailItem := AMail;
end;

procedure TAddInModule.AssignOLAttachmentToIdMsg(AOLMailItem: MailItem;
  out AIdMsg: TIdMessage);
var
  i: integer;
  LPath, LFileName: string;
begin
  LPath := 'c:\temp\';

  for i := 1 to AOLMailItem.Attachments.Count do
  begin
    LFileName := LPath + AOLMailItem.Attachments.Item(i).FileName;
    AOLMailItem.Attachments.Item(i).SaveAsFile(LFileName);
    TIdAttachmentFile.Create(AIdMsg.MessageParts, LFileName);
  end;
end;

function TAddInModule.AssignOLMailItemToIdMessage(AOLMailItem: MailItem;
  out AIdMsg: TIdMessage): boolean;
begin
  Result := False;
  AIdMsg.From.Address := AOLMailItem.SenderEmailAddress;
  AIdMsg.Subject := AOLMailItem.Subject;
  AIdMsg.Body.Text := AOLMailItem.Body;
  ShowMessage(AOLMailItem.Subject + ':' + AOLMailItem.Body);// IntToStr(AOLMailItem.Attachments.Count));
  AssignOLRecipientToIdMsg(AOLMailItem, AIdMsg);
  AssignOLAttachmentToIdMsg(AOLMailItem, AIdMsg);
  Result := True;
end;

procedure TAddInModule.AssignOLRecipientToIdMsg(AOLMailItem: MailItem;
  out AIdMsg: TIdMessage);
var
  i: integer;
  LItem: TIdEMailAddressItem;
begin
  for i := 1 to AOLMailItem.Recipients.Count do
  begin
    case AOLMailItem.Recipients.Item(i).type_ of
      //olTo
      1: begin
        LItem := AIdMsg.Recipients.Add;
      end;
      //olCC
      2: begin
        LItem := AIdMsg.CCList.Add;
      end;
      //olBCC
      3: begin
        LItem := AIdMsg.BCCList.Add;
      end;
      else
        LItem := nil;
    end;

    if Assigned(LItem) then
    begin
      LItem.Address := AOLMailItem.Recipients.Item(i).Address;
    end;
  end;
end;

procedure TAddInModule.AsyncEmailDisplay;
begin
  Parallel.Async(
    procedure (const task: IOmniTask)
    var
      i: integer;
      handles: array [0..1] of THandle;
      msg    : TOmniMessage;
      rec    : TEntryIdRecord;
      LID: TID;
    begin
      handles[0] := FEmailDisplayStopEvent.Handle;
      handles[1] := FEmailDisplayMQ.GetNewMessageEvent;

      while WaitForMultipleObjects(2, @handles, false, INFINITE) = (WAIT_OBJECT_0 + 1) do
      begin
        while FEmailDisplayMQ.TryDequeue(msg) do
        begin
          rec := msg.MsgData.ToRecord<TEntryIdRecord>;
          if msg.MsgID = 1 then //Request Mail View
          begin
            task.Invoke(
              procedure
              begin
                if (rec.FEntryID <> '') and (rec.FStoreID <> '') then
                  ShowMailContents(rec.FEntryID, rec.FStoreID);
              end
            );
          end
          else
          if msg.MsgID = 2 then //Request Reply Mail
          begin
            task.Invoke(
              procedure
              begin
                if (rec.FEntryID <> '') and (rec.FStoreID <> '') then
                  ReplyMail(rec);
              end
            );
          end
          else
          if msg.MsgID = 3 then //Request Create Mail
          begin
            task.Invoke(
              procedure
              begin
                if (rec.FTo <> '') and (rec.FSubject <> '') then
                  CreateMail(rec);
              end
            );
          end
          else
          if msg.MsgID = 4 then //Request Forward Mail
          begin
            task.Invoke(
              procedure
              begin
                if (rec.FEntryID <> '') and (rec.FStoreID <> '') then
                  ForwardMail(rec);
              end
            );
          end
          else
          if msg.MsgID = 5 then //Request Mail View From .msg File
          begin
            task.Invoke(
              procedure
              begin
                if FileExists(rec.FAttachFileName) then
                  ViewMailFromMsgFile(rec.FAttachFileName);
              end
            );
          end
        end;
      end;
    end
  );
end;

procedure TAddInModule.AsyncSendEntryId2IPC(AIsDrag: Boolean);
var
  LMailItem: MailItem;
  LFolder: MAPIFolder;
  LRec: TOLMsgFileRecord;
  LOmniValue: TOmniValue;
begin
  Parallel.Async(
    procedure (const task: IOmniTask)
    var
      i: integer;
    begin
      i := OutlookApp.ActiveExplorer.Selection.Count;

      if (i > 1) and (AIsDrag) then
      begin
        ShowMessage('DragDrop 기능을 이용하기 위해서는 메일을 1개만 선택 하세요');
        exit;
      end;

      LFolder := OutlookApp.ActiveExplorer.CurrentFolder as MAPIFolder;

        LMailItem := OutlookApp.ActiveExplorer.Selection.Item(i) as MailItem;

        LRec.FEntryId := LMailItem.EntryID;
        LRec.FStoreId := LFolder.StoreID;
        LRec.FSender := LMailItem.SenderEmailAddress;
        LRec.FReceiveDate := LMailItem.ReceivedTime;
        LRec.FSubject := LMailItem.Subject;
        LRec.FMailItem := LMailItem;
        LRec.FUserEmail := FUserEmail;
        LRec.FUserName := FUserName;

        LOmniValue := TOmniValue.FromRecord<TOLMsgFileRecord>(LRec);
        FOLMsg2IPCMQ.Enqueue(TOmniMessage.Create(2, LOmniValue));
    end );
end;

procedure TAddInModule.CreateAppointment(ATodoItem: variant);
var
  LAppointment: AppointmentItem;
  LStr: string;
  LTimeLog: TTimeLog;
begin
  LAppointment := OutlookApp.CreateItem(olAppointmentItem) as AppointmentItem;

  if Assigned(LAppointment) then
  begin
    LTimeLog := ATodoItem.Start;
    LAppointment.Start := TimeLogToDateTime(LTimeLog);
    LTimeLog := ATodoItem.End_;
    LAppointment.End_ := TimeLogToDateTime(LTimeLog);
//    LAppointment.Location;
    LAppointment.Body := ATodoItem.Body;
//    LAppointment.AllDayEvent := ATodoItem.Subject;
    LAppointment.Subject := ATodoItem.Subject;
    LAppointment.Save;
//    LAppointment.Display(True);
  end;
end;

procedure TAddInModule.CreateHttpServer4WS(APort, ATransmissionKey: string;
  aClient: TInterfacedClass; const aInterfaces: array of TGUID);
begin
  if not Assigned(FRestServer) then
  begin
    // initialize a TObjectList-based database engine
    FRestServer := TSQLRestServerFullMemory.CreateWithOwnModel([]);
    // register our Interface service on the server side
    FRestServer.CreateMissingTables;
    FServiceFactoryServer := FRestServer.ServiceDefine(aClient, aInterfaces , sicShared) as TServiceFactoryServer;
    FServiceFactoryServer.SetOptions([], [optExecLockedPerInterface]). // thread-safe fConnected[]
      ByPassAuthentication := true;

//    FRestMode := rmWebSocket;

//    FRestServer.OnSessionCreate := SessionCreate;
//    FRestServer.OnSessionClosed := SessionClosed;
  end;

  if not Assigned(FHTTPServer) then
  begin
    // launch the HTTP server
    FPortName := APort;
    FHTTPServer := TSQLHttpServer.Create(APort, [FRestServer], '+' , useBidirSocket);
    FHTTPServer.WebSocketsEnable(FRestServer, ATransmissionKey);
    FIsServerActive := True;
  end;

  FCommModes := FCommModes + [cmWebSocket];
end;

procedure TAddInModule.CreateMail(AEntryIdRecord: TEntryIdRecord);
var
  LMailItem: MailItem;
  LStr: string;
begin
  LMailItem := OutlookApp.CreateItem(olMailItem) as MailItem;
  if Assigned(LMailItem) then
  begin
    LMailItem.To_ := AEntryIdRecord.FTo;
    LMailItem.Subject := AEntryIdRecord.FSubject;
    LStr := Utf8ToString(Base64ToBin(StringToUtf8(AEntryIdRecord.FHTMLBody)));
    LMailItem.HTMLBody := LStr + LMailItem.HTMLBody;

    if AEntryIdRecord.FAttached <> '' then
    begin
      LStr := ExtractFileName(AEntryIdRecord.FAttachFileName);
      FileFromString(AEntryIdRecord.FAttached, AEntryIdRecord.FAttachFileName);
      LMailItem.Attachments.Add(AEntryIdRecord.FAttachFileName,olByValue,1,LStr);
    end;

    LMailItem.Display(False);
    System.SysUtils.DeleteFile(AEntryIdRecord.FAttachFileName);
  end;
end;

procedure TAddInModule.DeleteMsgFilesFromTempFolder;
//var
//  ShOp: TSHFileOpStruct;
begin
//  ShOp.Wnd := Self.Handle;
//  ShOp.wFunc := FO_DELETE;
//  ShOp.pFrom := PChar(FOLDER_NAME_4_TEMP_MSG_FILES);
//  ShOp.pTo := nil;
//  ShOp.fFlags := FOF_NO_UI;
//  SHFileOperation(ShOp);
  DirectoryDelete(FOLDER_NAME_4_TEMP_MSG_FILES);
end;

procedure TAddInModule.DestroyHttpServer;
begin
  if Assigned(FHTTPServer) then
    FreeAndNil(FHTTPServer);

  if Assigned(FRestServer) then
  begin
    FRestServer := nil
  end;

  if Assigned(FModel) then
    FreeAndNil(FModel);
end;

procedure TAddInModule.DoItemAdd(ASender: TObject; const Item: IDispatch);
begin

end;

procedure TAddInModule.ForwardMail(AEntryIdRecord: TEntryIdRecord);
var
  LMailItem,
  LReplyMail: MailItem;
  LStr: string;
begin
  LMailItem := FNameSpace.GetItemFromID(AEntryIdRecord.FEntryId,
    AEntryIdRecord.FStoreId) as MailItem;

  if Assigned(LMailItem) then
  begin
    LReplyMail := LMailItem.Forward;
    LStr := Utf8ToString(Base64ToBin(StringToUtf8(AEntryIdRecord.FHTMLBody)));
    LReplyMail.HTMLBody := LStr + LReplyMail.HTMLBody;

    if AEntryIdRecord.FAttached <> '' then
    begin
      LStr := ExtractFileName(AEntryIdRecord.FAttachFileName);
      FileFromString(AEntryIdRecord.FAttached, AEntryIdRecord.FAttachFileName);
      LReplyMail.Attachments.Add(AEntryIdRecord.FAttachFileName,olByValue,1,LStr);
    end;

    LReplyMail.Display(False);
    System.SysUtils.DeleteFile(AEntryIdRecord.FAttachFileName);
  end;
end;

procedure TAddInModule.GetAutoForwardFolderAccounList;
begin
//  FAutoForwardFolderPathList.Add('\\jhpark@hyundai-gs.com\받은 편지함');
end;

function TAddInModule.GetEmail2StrList: TStringList;
var
  LMailItem: MailItem;
  i,j: integer;
  LFolder: MAPIFolder;
  LReceiver, LCC, LBCC: string;
  Docs: TVariantDynArray;
  DocsDA: TDynArray;
  LCount: integer;
begin
  Result := TStringList.Create;

  LCount := OutlookApp.ActiveExplorer.Selection.Count;

  if LCount = 0 then
  begin
    ShowMessage('DragDrop 기능을 이용하기 위해서는 메일을 1개 이상 선택 하세요');
    exit;
  end;

  DocsDA.Init(TypeInfo(TVariantDynArray), Docs, @LCount);
  LCount := OutlookApp.ActiveExplorer.Selection.Count;
  SetLength(Docs,LCount);

  Result.Add('ServerName='+IPC_SERVER_NAME_4_INQMANAGE);
  Result.Add('Command='+CMD_SEND_MAIL_ENTRYID2);

  for i := 1 to OutlookApp.ActiveExplorer.Selection.Count do
  begin
    LMailItem := OutlookApp.ActiveExplorer.Selection.Item(i) as MailItem;
    LFolder := OutlookApp.ActiveExplorer.CurrentFolder as MAPIFolder;

    for j := 1 to LMailItem.Recipients.Count do
      if LMailItem.Recipients.Item(j).type_ = OlTo then
        LReceiver := LReceiver + LMailItem.Recipients.Item(j).Address + ';'
      else if LMailItem.Recipients.Item(j).type_ = olCC then
        LCC := LCC + LMailItem.Recipients.Item(j).Address + ';'
      else if LMailItem.Recipients.Item(j).type_ = olBCC then
        LBCC := LBCC + LMailItem.Recipients.Item(j).Address + ';';

    TDocVariant.New(Docs[i-1]);

    Docs[i-1].EntryId := LMailItem.EntryID;
    Docs[i-1].StoreId := LFolder.StoreID;
    Docs[i-1].Sender := LMailItem.SenderEmailAddress;
    Docs[i-1].Receiver := LReceiver;
    Docs[i-1].RecvDate := DateTimeToStr(LMailItem.ReceivedTime);
    Docs[i-1].CC := LCC;
    Docs[i-1].BCC := LBCC;
    Docs[i-1].Subject := LMailItem.Subject;
    Docs[i-1].FolderPath := LFolder.FullFolderPath;
  end;

  LReceiver := Utf8ToString(DocsDA.SaveToJson);
//  ShowMessage(LReceiver);
  Result.Add('MailInfos='+LReceiver);
end;

function TAddInModule.GetFolderFromPath(APath: string): MAPIFOLDER;
var
  LStrArr: System.Types.TStringDynArray;
  LPath: string;
  i: integer;
  LFoundFolder : MAPIFOLDER;
  LSubFolders: _Folders;
begin
  LPath := StringReplace(APath, '\\', '', [rfReplaceAll]);
  LStrArr := SplitString(LPath, '\');

  LFoundFolder := OutlookApp.Session.Folders.Item(LStrArr[0]) as MAPIFOLDER;

  for i := 1 to High(LStrArr) do
  begin
    LSubFolders := LFoundFolder.Folders;
    LFoundFolder := LSubFolders.Item(LStrArr[i]) as MAPIFOLDER
  end;

  Result := LFoundFolder;
end;

procedure TAddInModule.GetInboxList;
var
  i,j,k: integer;
  LFolders: _Folders;
  LFolderName: string;
begin
  FNameSpace := OutlookApp.GetNamespace('MAPI') as _NameSpace;//.GetDefaultFolder();
  FUserEmail := FNameSpace.Accounts.Item(1).DisplayName;
  FUserName := FNameSpace.Accounts.Item(1).UserName;
  FIsPJHPC := FUserEmail = OLMYMAILADDR;

  for j := 1 to FNameSpace.Folders.Count do
  begin
    LFolders := FNameSpace.Folders.Item(j).Folders;
    LFolderName := FNameSpace.Folders.Item(j).Name;

    for k := 1 to LFolders.Count do
    begin
      if LFolders.Item(k).Name = '받은 편지함' then
      begin
        FInboxStoreIdList.Add(LFolderName + '=' + LFolders.Item(k).StoreID);
      end
      else
      if LFolders.Item(k).Name = 'Working' then
      begin
        g_WorkingFolder := LFolders.Item(k);
      end;
    end;
  end;
end;

function TAddInModule.GetOLEmailAccountInfo: RawUTF8;
var
  LOLAccountInfo: TOLAccountInfo;
begin
  LOLAccountInfo.SmtpAddress := FNameSpace.Accounts.Item(1).SmtpAddress;
  LOLAccountInfo.DisplayName := FNameSpace.Accounts.Item(1).DisplayName;
  LOLAccountInfo.UserName := FNameSpace.Accounts.Item(1).UserName;

  Result := RecordSaveJson(LOLAccountInfo, TypeInfo(TOLAccountInfo));
end;

function TAddInModule.GetResponse4MoveFolder2StrList(
  AEntryIdRecord: TEntryIdRecord): TStringList;
begin
  Result := TStringList.Create;

  Result.Add('ServerName='+IPC_SERVER_NAME_4_INQMANAGE);
  Result.Add('Command='+CMD_RESPONDE_MOVE_FOLDER_MAIL);
  Result.Add('NewEntryId='+AEntryIdRecord.FNewEntryId);
  Result.Add('MovedStoreId='+AEntryIdRecord.FStoreId4Move);
  Result.Add('MovedFolderPath='+AEntryIdRecord.FFolderPath);
end;

procedure TAddInModule.InitNetwork;
begin
  CreateHttpServer4WS(OL_PORT_NAME_4_WS, OL4WS_TRANSMISSION_KEY, TServiceOL4WS, [IOLMailService]);
end;

function TAddInModule.IsExistFolder(var AFolder: MAPIFolder;
  AFolderName: string): Boolean;
var
  i: integer;
begin
  Result := False;

  for i := 1 to AFolder.Folders.Count do
  begin
    if AFolder.Folders.Item(i).Name = AFoldername then
    begin
      AFolder := AFolder.Folders.Item(i);
      Result := True;
      break;
    end;
  end;
end;

procedure TAddInModule.MoveMail2Folder(var AEntryIdRecord: TEntryIdRecord);
var
  LMailItem,
  LMailItem2 : MailItem;
  LFolder: MAPIFolder;
begin
  LMailItem := FNameSpace.GetItemFromID(AEntryIdRecord.FEntryId,
    AEntryIdRecord.FStoreId) as MailItem;

  if Assigned(LMailItem) then
  begin
    LFolder := GetFolderFromPath(AEntryIdRecord.FFolderPath);

    if AEntryIdRecord.FIsCreateHullNoFolder then
    begin
      if AEntryIdRecord.FHullNo <> '' then
      begin
        if not IsExistFolder(LFolder, AEntryIdRecord.FHullNo) then
          LFolder := LFolder.Folders.Add(AEntryIdRecord.FHullNo, olFolderInbox);
      end;

      if AEntryIdRecord.FSubFolder <> '' then
      begin
        if not IsExistFolder(LFolder, AEntryIdRecord.FSubFolder) then
          LFolder := LFolder.Folders.Add(AEntryIdRecord.FSubFolder, olFolderInbox);
      end;
    end;

    LMailItem2 := LMailItem.Move(LFolder) as MailItem;
    AEntryIdRecord.FNewEntryId := LMailItem2.EntryID;
    AEntryIdRecord.FFolderPath := LFolder.FullFolderPath;
  end;
end;

procedure TAddInModule.MoveMail2WorkingFolder(AEntryIDList: WideString);
const
  GS_EMAIL = 'hyundai-gs.com';
var
  IFolderInbox: MAPIFolder;
  LNameSpace: _NameSpace;
  LMailItem: MailItem;
  LAccount: _Account;
  LStrArr: System.Types.TStringDynArray;
  LStoreID: string;
  LRec: TOLMsgFileRecord;
  LOmniValue: TOmniValue;
  i,j,k: integer;
begin
  LNameSpace := OutlookApp.GetNamespace('MAPI') as _NameSpace;
  LStrArr := SplitString(AEntryIDList, ',');
  Log4OL('받은 편지함!', True);
  for i := Low(LStrArr) to High(LStrArr) do
  begin
    LMailItem := nil;

    for k := 0 to FInboxStoreIDList.Count - 1 do
    begin
      LStoreID := FInboxStoreIDList.ValueFromIndex[k];

      try
        LMailItem := LNameSpace.GetItemFromID(LStrArr[i],LStoreID) as MailItem;

        if Assigned(LMailItem) then
        begin
          AssignMailItem2Rec(LMailItem, LRec);
          LOmniValue := TOmniValue.FromRecord<TOLMsgFileRecord>(LRec);
          if not FOLMsgQueue.Enqueue(TOmniMessage.Create(1, LOmniValue)) then
            Log4OL('Send queue is full!', True)
          else
            Log4OL('Send queue is success!', True);

          Log4OL('Receiver : ' + LRec.FReceiver, True);
          Break;
        end;
      except
        continue;
      end;
    end;//for
  end;//for
end;

function TAddInModule.ProcessCommandFromClient(ACommand: string): RawUTF8;
var
  Command: string;
  LStrList,
  LStrList2: TStringList;
  LMailItem: MailItem;
  LEntryId, LStoreId, LAttached: string;
  rec    : TEntryIdRecord;
  LOmniValue: TOmniValue;
  i: integer;
  LFolder: MAPIFolder;
  LRec2: TOLMsgFileRecord;
begin
  LStrList := TStringList.Create;
  try
    LStrList.Text := ACommand;
    Command := LStrList.Values['Command'];
    Log4OL(CMD_REQ_MAIL_VIEW, True);

    if Command = CMD_REQ_MAIL_VIEW then
    begin
      LEntryId := LStrList.Values['EntryId'];
      LStoreId := LStrList.Values['StoreId'];
      rec.FEntryId := LEntryId;
      rec.FStoreId := LStoreId;

      LOmniValue := TOmniValue.FromRecord<TEntryIdRecord>(rec);
      FEmailDisplayMQ.Enqueue(TOmniMessage.Create(1, LOmniValue));
    end
    else
    if Command = CMD_REQ_REPLY_MAIL then
    begin
      rec.FEntryId := LStrList.Values['EntryId'];;
      rec.FStoreId := LStrList.Values['StoreId'];;
      rec.FHTMLBody := LStrList.Values['HTMLBody'];

      LAttached := LStrList.Values['TaskInfoAttached'];
      if LAttached <> '' then
      begin
        rec.FAttached := LAttached;
        rec.FAttachFileName := LStrList.Values['AttachedFileName'];
      end;

      LOmniValue := TOmniValue.FromRecord<TEntryIdRecord>(rec);
      FEmailDisplayMQ.Enqueue(TOmniMessage.Create(2, LOmniValue));
    end
    else
    if Command = CMD_REQ_CREATE_MAIL then
    begin
      rec.FEntryId := LStrList.Values['EntryId'];;
      rec.FStoreId := LStrList.Values['StoreId'];;
      rec.FSubject := LStrList.Values['Subject'];
      rec.FTo := LStrList.Values['To'];
      rec.FHTMLBody := LStrList.Values['HTMLBody'];

      LAttached := LStrList.Values['TaskInfoAttached'];
      if LAttached <> '' then
      begin
        rec.FAttached := LAttached;
        rec.FAttachFileName := LStrList.Values['AttachedFileName'];
      end;

      LOmniValue := TOmniValue.FromRecord<TEntryIdRecord>(rec);
      FEmailDisplayMQ.Enqueue(TOmniMessage.Create(3, LOmniValue));
    end
    else
    if Command = CMD_REQ_FORWARD_MAIL then
    begin
      rec.FEntryId := LStrList.Values['EntryId'];;
      rec.FStoreId := LStrList.Values['StoreId'];;
      rec.FHTMLBody := LStrList.Values['HTMLBody'];

      LAttached := LStrList.Values['TaskInfoAttached'];
      if LAttached <> '' then
      begin
        rec.FAttached := LAttached;
        rec.FAttachFileName := LStrList.Values['AttachedFileName'];
      end;

      LOmniValue := TOmniValue.FromRecord<TEntryIdRecord>(rec);
      FEmailDisplayMQ.Enqueue(TOmniMessage.Create(4, LOmniValue));
    end
    else
    if Command = CMD_REQ_MAILINFO_SEND then
    begin
      i := OutlookApp.ActiveExplorer.Selection.Count;

      if i > 1 then
      begin
        ShowMessage('DragDrop 기능을 이용하기 위해서는 메일을 1개만 선택 하세요');
        exit;
      end;

      LFolder := OutlookApp.ActiveExplorer.CurrentFolder as MAPIFolder;

      LMailItem := OutlookApp.ActiveExplorer.Selection.Item(i) as MailItem;

      LRec2.FEntryId := LMailItem.EntryID;
      LRec2.FStoreId := LFolder.StoreID;
      LRec2.FSender := LMailItem.SenderEmailAddress;
      LRec2.FReceiveDate := LMailItem.ReceivedTime;
      LRec2.FSubject := LMailItem.Subject;
      LRec2.FMailItem := LMailItem;
      LRec2.FUserEmail := FUserEmail;
      LRec2.FUserName := FUserName;
      LRec2.FSavedOLFolderPath := LFolder.FullFolderPath;

      LStrList.Clear;
      LStrList.Add('Command='+CMD_SEND_MAIL_ENTRYID);
      LStrList.Add('EntryId='+LRec2.FEntryId);
      LStrList.Add('StoreId='+LRec2.FStoreId);
      LStrList.Add('Sender='+LRec2.FSender);
      LStrList.Add('Receiver='+LRec2.FReceiver);
      LStrList.Add('RecvDate='+DateTimeToStr(LRec2.FReceiveDate));
      LStrList.Add('CC='+LRec2.FCarbonCopy);
      LStrList.Add('BCC='+LRec2.FBlindCC);
      LStrList.Add('Subject='+LRec2.FSubject);
      LStrList.Add('FolderPath='+LRec2.FSavedOLFolderPath);
      LStrList.Add('UserEmail='+LRec2.FUserEmail);
      LStrList.Add('UserName='+LRec2.FUserName);
      Result := LStrList.Text;
    end
    else
    if Command = CMD_REQ_MAIL_VIEW_FROM_MSGFILE then
    begin
      LEntryId := LStrList.Values['FileName'];
      rec.FAttachFileName := LEntryId;

      LOmniValue := TOmniValue.FromRecord<TEntryIdRecord>(rec);
      FEmailDisplayMQ.Enqueue(TOmniMessage.Create(5, LOmniValue));
    end;
  finally
    LStrList.Free;
  end;
end;

procedure TAddInModule.ReplyMail(AEntryIdRecord: TEntryIdRecord);
var
  LMailItem,
  LReplyMail: MailItem;
  LStr: string;
begin
  LMailItem := FNameSpace.GetItemFromID(AEntryIdRecord.FEntryId,
    AEntryIdRecord.FStoreId) as MailItem;

  if Assigned(LMailItem) then
  begin
    LReplyMail := LMailItem.Reply;
//    LReplyMail.HTMLBody := AEntryIdRecord.FHTMLBody;
    LStr := Utf8ToString(Base64ToBin(StringToUtf8(AEntryIdRecord.FHTMLBody)));
    LReplyMail.HTMLBody := LStr + LReplyMail.HTMLBody;

    if AEntryIdRecord.FAttached <> '' then
    begin
      LStr := ExtractFileName(AEntryIdRecord.FAttachFileName);
      FileFromString(AEntryIdRecord.FAttached, AEntryIdRecord.FAttachFileName);
      LReplyMail.Attachments.Add(AEntryIdRecord.FAttachFileName,olByValue,1,LStr);
    end;

    LReplyMail.Display(False);
    System.SysUtils.DeleteFile(AEntryIdRecord.FAttachFileName);
  end;
end;

procedure TAddInModule.SendEntryID2IPCFromList;
var
  LOmniValue: TOmniValue;
  i: integer;
begin
//  if FIPCClientList.Count = 0 then
//    exit;

  for i := FEntryIDList.Count - 1 downto 0 do
  begin
    LOmniValue := FEntryIDList.Strings[i];

    if not FOLMsgQueue.Enqueue(TOmniMessage.Create(1, LOmniValue)) then
      Log4OL('Send queue is full!', True)
    else
    begin
      FEntryIDList.Delete(i);
      Log4OL('Send queue is success!', True);
    end;
  end;
end;

procedure TAddInModule.SendMailToMsgFile(AEntryIDList: WideString);
const
  GS_EMAIL = 'hyundai-gs.com';
var
  LMailItem: MailItem;
  LStrArr: System.Types.TStringDynArray;
  i,j,k: integer;
  LFolders: _Folders;
  LStrGuid: string;
  LStrFile: string;
  LRec: TOLMsgFile4STOMP;
  LOmniValue: TOmniValue;
  LRaw: RawByteString;
  LUtf8: RawUTF8;
begin
  LStrArr := SplitString(AEntryIDList, ',');

  for i := Low(LStrArr) to High(LStrArr) do
  begin
    LMailItem := nil;

    for j := 1 to FNameSpace.Folders.Count do
    begin
        LFolders := FNameSpace.Folders.Item(j).Folders;

        for k := 1 to LFolders.Count do
        begin
//          if LFolders.Item(k).Name = '받은 편지함' then
//          if FAutoForwardFolderPathList.IndexOf(LFolders.Item(k).FolderPath) <> -1 then
//          begin
            try
              LMailItem := FNameSpace.GetItemFromID(LStrArr[i],LFolders.Item(k).StoreID) as MailItem;
            except
              continue;
            end;

            if Assigned(LMailItem) then
            begin
              LStrGuid := EnsureDirectoryExists('c:\temp\') +
                TGuid.NewGuid.ToString + '.msg';
              LMailItem.SaveAs(LStrGuid, olMSGUnicode);
              LRaw := StringFromFile(LStrGuid);
              LRaw := SynLZCompress(LRaw);
              LUtf8 := BinToBase64(LRaw);
              LStrFile := UTF8ToString(LUtf8);
              LRec.FHost := MQ_SERVER_IP;
              LRec.FUserId := MQ_USER_ID;
              LRec.FPasswd := MQ_PASSWORD;
              LRec.FMsgFile := LStrFile;
              LOmniValue := TOmniValue.FromRecord<TOLMsgFile4STOMP>(LRec);

              if not FOLMsg2STOMPMQ.Enqueue(TOmniMessage.Create(1, LOmniValue)) then
                Log4OL('Send queue is full!', True)
              else
                Log4OL('Send queue is success!', True);

              SysUtils.DeleteFile(LStrGuid);
            end;
//          end;
        end;//for
    end;
  end;
end;

procedure TAddInModule.SendMailToMsgFile_Async(AEntryIDList: WideString);
begin
  Parallel.Async(
    procedure (const task: IOmniTask)
    begin
      SendMailToMsgFile(AEntryIDList);
    end );
end;

procedure TAddInModule.Send_Message(AHead, ATitle, AContent, ASendUser,
  ARecvUser, AFlag: string);
var
  lstr,
  lcontent : String;
begin
//  헤더의 길이가 21byte를 넘지 않아야 함.
//  lhead := 'HiTEMS-문제점보고서';
//  ltitle   := '업무변경건';
  lcontent := AContent;

  if Aflag = 'B' then
  begin
    while True do
    begin
      if lcontent = '' then
        Break;

      if Length(lcontent) > 90 then
      begin
        lstr := Copy(lcontent,1,90);
        lcontent := Copy(lcontent,91,Length(lcontent)-90);
      end else
      begin
        lstr := Copy(lcontent,1,Length(lcontent));
        lcontent := '';
      end;

      //문자 메세지는 title(lstr)만 보낸다.
//      Send_Message_Main_CODE(AFlag,ASendUser,ARecvUser,AHead,lstr,ATitle);
    end;
  end
  else
  begin
    lstr := lcontent;
//    Send_Message_Main_CODE(AFlag,ASendUser,ARecvUser,AHead,lstr,ATitle);
  end;
end;

function TAddInModule.ServerExecuteFromClient(ACommand: string): RawUTF8;
var
  LEntryId, LStoreId: string;
  rec    : TEntryIdRecord;
  LOmniValue: TOmniValue;
  LMailItem: MailItem;
  Command, LJson: String;
  LStrList,
  LStrList2: TStringList;
  LVarArr: TVariantDynArray;
  i: integer;
begin
  Result := '';
  LStrList := TStringList.Create;
  try
    LStrList.Text := ACommand;
    Command := LStrList.Values['Command'];

    if Command = CMD_REQ_MAILINFO_SEND2 then
    begin
      LStrList2 := GetEmail2StrList;
      try
        Result := LStrList2.Text;
      finally
        LStrList2.Free;
      end;
    end
    else
    if Command = CMD_REQ_MOVE_FOLDER_MAIL then
    begin
      LEntryId := LStrList.Values['EntryId'];
      LStoreId := LStrList.Values['StoreId'];
      rec.FEntryId := LEntryId;
      rec.FStoreId := LStoreId;
      rec.FStoreId4Move := LStrList.Values['MoveStoreId'];
      rec.FFolderPath := LStrList.Values['MoveStorePath'];
      rec.FHullNo := LStrList.Values['HullNo'];
      rec.FSubFolder := LStrList.Values['SubFolderName'];
      rec.FIsCreateHullNoFolder := StrToBool(LStrList.Values['IsCreateHullNoFolder']);

      MoveMail2Folder(rec);
      LStrList2 := GetResponse4MoveFolder2StrList(rec);
      try
        Result := LStrList2.Text;
      finally
        LStrList2.Free;
      end;
    end
    else
    if Command = CMD_REQ_ADD_APPOINTMENT then
    begin
      LJson := LStrList.Values['TodoItemsJson'];
      LVarArr := JSONToVariantDynArray(LJson);

      for i := 0 to High(LVarArr) do
      begin
        CreateAppointment(LVarArr[i]);
      end;
    end
    else
    if Command = CMD_SEND_MAIL_2_MSGFILE then
    begin
      LEntryId := LStrList.Values['EntryId'];
      LStoreId := LStrList.Values['StoreId'];
      LMailItem := FNameSpace.GetItemFromID(LEntryId,LStoreId) as MailItem;

      if Assigned(LMailItem) then
      begin
        LEntryId := EnsureDirectoryExists(FOLDER_NAME_4_TEMP_MSG_FILES) +
          TGuid.NewGuid.ToString + '.msg';
        LMailItem.SaveAs(LEntryId, olMSGUnicode);
        Result := StringToUTF8(LEntryId);
      end;
    end;
  finally
    LStrList.Free;
  end;
end;

function TAddInModule.SessionClosed(Sender: TSQLRestServer;
  Session: TAuthSession; Ctxt: TSQLRestServerURIContext): boolean;
begin
//  DeleteConnectionFromLV(Session.RemoteIP, FPortName, Session.ID, Session.User.LogonName);
//  Result := False;
end;

function TAddInModule.SessionCreate(Sender: TSQLRestServer;
  Session: TAuthSession; Ctxt: TSQLRestServerURIContext): boolean;
begin
//  AddConnectionToLV(Session.RemoteIP, FPortName, Session.ID, Session.User.LogonName);
//  Result := False;
end;

procedure TAddInModule.ShowMailContents(AEntryId, AStoreId: string);
var
  LMailItem: MailItem;
  LFolder: MAPIFolder;
begin
  LMailItem := FNameSpace.GetItemFromID(AEntryId, AStoreId) as MailItem;

  if Assigned(LMailItem) then
  begin
    LMailItem.Display(False);
  end;
end;

procedure TAddInModule.ShowStoreIdFromSelected;
var
  i: integer;
  LFolder: MAPIFolder;
  LStr: string;
begin
  i := OutlookApp.ActiveExplorer.Selection.Count;

  if i > 1 then
  begin
    ShowMessage('이 기능을 이용하기 위해서는 메일을 1개만 선택 하세요');
    exit;
  end;

  LFolder := OutlookApp.ActiveExplorer.CurrentFolder as MAPIFolder;

  if Assigned(LFolder) then
  begin
    LStr := LFolder.Name + #13#10 + LFolder.FullFolderPath + #13#10 +LFolder.StoreID;
    ShowMessage(LStr);
    Clipboard.AsText := LFolder.FullFolderPath + '=' + LFolder.StoreID;
  end;
end;

procedure TAddInModule.Timer1Timer(Sender: TObject);
begin
  Timer1.Enabled := False;

  GetAutoForwardFolderAccounList;
  GetInboxList;
  Log4OL('GetInboxList() Executed', True);
  FWorker4OLMsg := TWorker4OLMsg.Create(FOLMsgQueue, FOLMsg2IPCMQ,
    FInboxStoreIdList, FNameSpace);
  FWorker4STOMP := TWorker4STOMP.Create(FOLMsg2STOMPMQ);//FIPCClientList

  Log4OL('FEntryIDList Created', True);
  AdjustDisplayMenu;
  Log4OL('AdjustDisplayMenu() Executed', True);
  AsyncEmailDisplay;
  Log4OL('AsyncEmailDisplay() Executed', True);
  InitNetwork;
  Log4OL('InitNetwork() Executed', True);

  FInitCompleted := True;
end;

procedure TAddInModule.ViewMailFromMsgFile(AFileName: string);
var
  LMailItem: MailItem;
  LFolder: MAPIFolder;
begin
  LMailItem := FNameSpace.OpenSharedItem(AFileName) as MailItem;

  if Assigned(LMailItem) then
  begin
    LMailItem.Display(False);
  end;
end;

{ TWorker4OLMsg }

constructor TWorker4OLMsg.Create(sendQueue, IPCQueue: TOmniMessageQueue;
  AInboxStoreIdList: TStringList; ANameSpace: _NameSpace);
begin
  inherited Create;

  FreeOnTerminate := True;
  FOLMsgQueue4Worker := sendQueue;
  FOLMsg2IPCMQ4Worker := IPCQueue;
  FInboxStoreIdList := AInboxStoreIdList;
  FNameSpace_Worker4OLMsg := ANameSpace;
  FStopEvent := TEvent.Create;
end;

destructor TWorker4OLMsg.Destroy;
begin
  FreeAndNil(FStopEvent);

  inherited;
end;

procedure TWorker4OLMsg.Execute;
var
  handles: array [0..1] of THandle;
  msg    : TOmniMessage;
  rec    : TOLMsgFileRecord;
  LMailItem : MailItem;
  LOmniValue: TOmniValue;
  i,k: integer;
  LStrArr: System.Types.TStringDynArray;
  LEntryIDList: string;
  LStoreID: string;
  LCheckReceiverOK: Boolean;
  LEntryRec: TEntryIdRecord;

  procedure SendIPCMq(AStoreId: string);
  var
    j: integer;
  begin
    LMailItem := FNameSpace_Worker4OLMsg.GetItemFromID(LStrArr[i], AStoreId) as MailItem;

    if Assigned(LMailItem) then
    begin
      rec.Clear;

      for j := 1 to LMailItem.Recipients.Count do
        if LMailItem.Recipients.Item(j).type_ = OlTo then
          rec.FReceiver := rec.FReceiver + LMailItem.Recipients.Item(j).Address + ';'
        else if LMailItem.Recipients.Item(j).type_ = olCC then
          rec.FCarbonCopy := rec.FCarbonCopy + LMailItem.Recipients.Item(j).Address + ';'
        else if LMailItem.Recipients.Item(j).type_ = olBCC then
          rec.FBlindCC := rec.FBlindCC + LMailItem.Recipients.Item(j).Address + ';';

      if LEntryRec.FIgnoreReceiver2pjh then
        LCheckReceiverOK := True
      else
      begin
        LCheckReceiverOK := Pos(OLMYMAILADDR, rec.FReceiver) > 0;
      end;

      if LCheckReceiverOK then
      begin
        rec.FSender := LMailItem.SenderEmailAddress;
        rec.FReceiveDate := LMailItem.ReceivedTime;
        rec.FSubject := LMailItem.Subject;
        rec.FMailItem := LMailItem;

        if not LEntryRec.FIgnoreEmailMove2WorkFolder then
        begin
          rec.FMailItem := LMailItem.Move(g_WorkingFolder) as MailItem;
          rec.FStoreId := g_WorkingFolder.StoreID;
        end
        else
          rec.FStoreId := AStoreId;

        rec.FEntryId := LMailItem.EntryID;
        LOmniValue := TOmniValue.FromRecord<TOLMsgFileRecord>(rec);
        FOLMsg2IPCMQ4Worker.Enqueue(TOmniMessage.Create(2, LOmniValue));
      end;
    end;
  end;
begin
  handles[0] := FStopEvent.Handle;
  handles[1] := FOLMsgQueue4Worker.GetNewMessageEvent;

  while WaitForMultipleObjects(2, @handles, false, INFINITE) = (WAIT_OBJECT_0 + 1) do
  begin
    if terminated then
      break;

    while FOLMsgQueue4Worker.TryDequeue(msg) do
    begin
      if msg.MsgID = 4 then //Send 2 IPC FolderID(Popup Menu)
      begin
        FOLMsg2IPCMQ4Worker.Enqueue(TOmniMessage.Create(3, msg.MsgData));
      end
      else
      begin
        LEntryRec := msg.MsgData.ToRecord<TEntryIdRecord>;
        LEntryIDList := LEntryRec.FEntryId;
        LStrArr := SplitString(LEntryIDList, ',');

        for i := Low(LStrArr) to High(LStrArr) do
        begin
          if LEntryRec.FStoreId <> '' then
          begin
            SendIPCMq(LEntryRec.FStoreId);
          end
          else
          begin
            for k := 0 to FInboxStoreIDList.Count - 1 do
            begin
              LStoreID := FInboxStoreIDList.ValueFromIndex[k];

              try
                SendIPCMq(LStoreID);
              except
                continue;
              end;
            end;//for
          end;//else
        end;//for
      end;//else
    end;//while
  end;//while
end;

procedure TWorker4OLMsg.Stop;
begin
  FStopEvent.SetEvent;
end;

{ TServiceOL4WS }

procedure TServiceOL4WS.CallbackReleased(const callback: IInvokable;
  const interfaceName: RawUTF8);
var
  LClientInfo: TClientInfo;
  LIndex,i: integer;
begin
  assert(interfaceName = 'IOLCallback');
  LIndex := -1;
  LIndex := InterfaceArrayDelete(fConnected, callback);

  if LIndex <> -1 then
  begin
    LClientInfo := TClientInfo(FClientInfoList.Objects[LIndex]);
    FClientInfoList.Delete(LIndex);
  end;
end;

constructor TServiceOL4WS.Create;
begin
  FClientInfoList :=  TStringList.Create;
end;

destructor TServiceOL4WS.Destroy;
var
  i: integer;
begin
  for i := 0 to FClientInfoList.Count - 1 do
    TClientInfo(FClientInfoList.Objects[i]).Free;

  FClientInfoList.Free;

  inherited;
end;

function TServiceOL4WS.GetOLEmailAccountInfo: RawUTF8;
begin
  Result := MyAddInModule.GetOLEmailAccountInfo;
end;

function TServiceOL4WS.GetOLEmailInfo(ACommand: string): RawUTF8;
begin
  Result := MyAddInModule.ProcessCommandFromClient(ACommand);
end;

procedure TServiceOL4WS.Join(const pseudo: string;
  const callback: IOLMailCallback);
begin
//  MyAddInModule.ServerExecuteFromClient(ACommand);
end;

function TServiceOL4WS.ServerExecute(const Acommand: string): RawUTF8;
begin
  Result := MyAddInModule.ServerExecuteFromClient(ACommand);
end;

{ TWorker4STOMP }

constructor TWorker4STOMP.Create(sendQueue: TOmniMessageQueue);
var
  LPath: string;
begin
  inherited Create;

  FreeOnTerminate := True;
  FWorker4STOMPQueue := sendQueue;
  FWorker4STOMPStopEvent := TEvent.Create;
  FAutoForwardFolderPathDic := TDictionary<string, MAPIFolder>.Create;
//  FAutoForwardFolderPathArray := VarArrayCreate([0, 0], varVariant);
  LPath := OLMYRECVFOLDERPATH;
  GetFolderPath2Dic(LPath);
end;

destructor TWorker4STOMP.Destroy;
begin
  FAutoForwardFolderPathDic.Clear;
  FAutoForwardFolderPathDic.Free;
  FreeAndNil(FWorker4STOMPStopEvent);

  inherited;
end;

procedure TWorker4STOMP.Execute;
var
  handles: array [0..1] of THandle;
  msg    : TOmniMessage;
  rec    : TOLMsgFile4STOMP;
  lClient: TStompClient;
  lStr   : WideString;
begin
  handles[0] := FWorker4STOMPStopEvent.Handle;
  handles[1] := FWorker4STOMPQueue.GetNewMessageEvent;
  lClient := TStompClient.Create;
  try
    while WaitForMultipleObjects(2, @handles, false, INFINITE) = (WAIT_OBJECT_0 + 1) do
    begin
      if terminated then
        break;

      while FWorker4STOMPQueue.TryDequeue(msg) do
      begin
        Log4OL('FWorker4STOMPQueue', True);
        rec := msg.MsgData.ToRecord<TOLMsgFile4STOMP>;
        lClient.SetUserName(rec.FUserId);
        lClient.SetPassword(rec.FPasswd);
        lStr := rec.FMsgFile;
        SendMailToMsgFileThread(lStr, lClient, rec.FHost);
//        lClient.Connect(rec.FHost, 61613, '', TStompAcceptProtocol.Ver_1_1);
//        try
//          try
//            lClient.Send(EMAIL_TOPIC_NAME, rec.FMsgFile);
//          finally
//            if lClient.Connected then
//              lClient.Disconnect;
//          end;
//        except
//        end;
      end;
    end;
  finally
    lClient.Free;
  end;
end;

procedure TWorker4STOMP.GetFolderPath2Dic(AFolderPath: string);
var
  LNameSpace : _NameSpace;
  LFolders: _Folders;
  i, j,k: integer;
begin
  LNameSpace := MyAddInModule.OutlookApp.GetNamespace('MAPI') as _NameSpace;
  for j := 1 to LNameSpace.Folders.Count do
  begin
    LFolders := LNameSpace.Folders.Item(j).Folders;

    for k := 1 to LFolders.Count do
    begin
      if LFolders.Item(k).FolderPath = AFolderPath then
      begin
        FAutoForwardFolderPathDic.Add(AFolderPath, LFolders.Item(k));
        exit;
      end;
    end;
  end;
end;

procedure TWorker4STOMP.SendMailToMsgFileThread(AEntryIDList: WideString;
  AStompClient: TStompClient; AHostAddr: string);
var
  LMailItem: MailItem;
  LStrArr: System.Types.TStringDynArray;
  i,j,k: integer;
//  LFolders: _Folders;
  LFolder: MAPIFolder;
  LStrGuid: string;
  LStrFile, LKey: string;
  LRec: TOLMsgFile4STOMP;
  LOmniValue: TOmniValue;
  LRaw: RawByteString;
  LUtf8: RawUTF8;
  LNameSpace : _NameSpace;
begin
  LNameSpace := MyAddInModule.OutlookApp.GetNamespace('MAPI') as _NameSpace;
  LStrArr := SplitString(AEntryIDList, ',');

  for i := Low(LStrArr) to High(LStrArr) do
  begin
    LMailItem := nil;

    try
      for LKey in FAutoForwardFolderPathDic.Keys do
      begin
        LFolder := FAutoForwardFolderPathDic.Items[LKey];
        try
          LMailItem := LNameSpace.GetItemFromID(LStrArr[i],LFolder.StoreID) as MailItem;

          Log4OL('FAutoForwardFolderPathDic.Keys', True);
          if Assigned(LMailItem) then
          begin
            LStrGuid := EnsureDirectoryExists('c:\temp\') +
              TGuid.NewGuid.ToString + '.msg';
            LMailItem.SaveAs(LStrGuid, olMSGUnicode);
            LRaw := StringFromFile(LStrGuid);
            LRaw := SynLZCompress(LRaw);
            LUtf8 := BinToBase64(LRaw);
            LStrFile := UTF8ToString(LUtf8);

            Log4OL(AHostAddr, True);
            AStompClient.Connect(AHostAddr, 61613, '', TStompAcceptProtocol.Ver_1_1);
            try
              try
                AStompClient.Send(EMAIL_TOPIC_NAME, LStrFile);
              finally
                if AStompClient.Connected then
                  AStompClient.Disconnect;
              end;
            except
            end;
            SysUtils.DeleteFile(LStrGuid);
          end;
        except
          continue;
        end;
      end;
    except
      continue;
    end;
  end;
end;

procedure TWorker4STOMP.Stop;
begin
  FWorker4STOMPStopEvent.SetEvent;
end;

initialization
  TadxFactory.Create(ComServer, TCoOLMail4InqManage, CLASS_CoOLMail4InqManage, TAddInModule);

end.
