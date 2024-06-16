unit UnitOLDataType;

interface

uses System.Classes, System.SysUtils, System.StrUtils, Vcl.StdCtrls,
{$IFDEF USE_OUTLOOK2016}
  Outlook2016_TLB,
{$ELSE}
  Outlook2010,
{$ENDIF}
  UnitEnumHelper
  ;

type
  //diolMailFromOL: IPC�� �̿��Ͽ� Outlook���κ��� ���ŵǴ� Mail Info
  //diolFolderFromOL: IPC�� �̿��Ͽ� Outlook���κ��� ���ŵǴ� Folder Info
  //dillFileFromDrag: Outlook �Ǵ� Ž����� ���� Drag�� ���� �Ǵ� ����
  TDataKindFromMQ = (dkmqNone, dkmqMailFromOL, dkmqFolderFromOL, dkmqFileFromDrag, dkmqFinal);

  TWSInfoRec = record
    FIPAddr,
    FPortNo,
    FTransKey,
    FServerName,
    FComputerName: string;
    FIsWSEnabled,
    FIsSendMQ,
    FNamedPipeEnabled,
    FIsRemoteMode: Boolean;
  end;

  TGUIDFileName = record
    HasInput: boolean;
    FileName: string[255];
  end;

  TOLMsgFile4STOMP = record
    FHost, FUserId, FPasswd, FTopic: string;
    FMsgFileName,
    FMsgFilePath,
    FMsgFile: string;
  end;

  TEntryIdRecord = record
    FEntryId,
    FStoreId,
    FStoreId4Move,
    FFolderPath,
    FNewEntryId,
    FSubject,
    FTo,
    FHTMLBody,
    FHullNo,
    FSubFolder,
    FAttached,
    FAttachFileName: string;
    FIgnoreReceiver2pjh: Boolean; //True = �����ڰ� pjh�ΰ� ������ ����
    FIgnoreEmailMove2WorkFolder: Boolean; //True = Working Folder�� �̵� ����
    //True = Move�ϰ��� ������ ���� �Ʒ��� HullNo Folder ���� �� ������ ������ ���� �̵� ��
    FIsCreateHullNoFolder: Boolean;
//    FIsShowMailContents: Boolean; //True = Mail Display
  end;

  TOLMsgFileRecord = record
    FEntryId,
    FStoreId,
    FSender,
    FReceiver,
    FCarbonCopy,
    FBlindCC,
    FSubject,
    FUserEmail,
    FUserName,
    FSavedOLFolderPath,
    FSpecialStatement: string;
    FMailItem: MailItem;
    FReceiveDate: TDateTime;
    FServiceType,
    FEmailKind: integer;

    procedure Clear;
  end;

  TOLAccountInfo = record
    SmtpAddress, DisplayName, UserName: string;
  end;

const
  IPC_SERVER_NAME_4_OUTLOOK = 'Mail2CromisIPCServer';
  //Response�� �ʿ��Ҷ� ���Ǵ� ������, �񵿱� ����� �ƴ�(�񵿱� ����� Response�� �ȵ�)
  IPC_SERVER_NAME_4_OUTLOOK2 = 'Mail2CromisIPCServer2';
  IPC_SERVER_NAME_4_INQMANAGE = 'Mail2CromisIPCClient';
  WS_SERVER_NAME_4_OUTLOOK = 'WSServer4OLMail';

  CMD_LIST = 'CommandList';
  CMD_SEND_MAIL_ENTRYID = 'Send Mail Entry Id';
  CMD_SEND_MAIL_ENTRYID2 = 'Send Mail Entry Id2';
  CMD_SEND_FOLDER_STOREID = 'Send Folder Store Id';
  CMD_SEND_MAIL_2_MSGFILE = 'Send Mail To Msg File';

  CMD_RESPONDE_MOVE_FOLDER_MAIL = 'Resonse for Move Mail to Folder';
  CMD_REQ_MAIL_VIEW = 'Request Mail View';
  CMD_REQ_MAIL_VIEW_FROM_MSGFILE = 'Request Mail View From .msg file';
  CMD_REQ_MAILINFO_SEND = 'Request Mail-Info to Send';
  //���ϸ���Ʈ���� ����, TaskID�� �ڵ����� ��
  CMD_REQ_MAILINFO_SEND2 = 'Request Mail-Info to Send2';
  CMD_REQ_MOVE_FOLDER_MAIL = 'Request Move Mail to Folder';
  CMD_REQ_REPLY_MAIL = 'Request Reply Mail';
  CMD_REQ_CREATE_MAIL = 'Request Create Mail';
  CMD_REQ_FORWARD_MAIL = 'Request Forward Mail';
  CMD_REQ_ADD_APPOINTMENT = 'Request Add Appointment';
  //Remote Command
  CMD_REQ_TASK_LIST = 'Request Task List';
  CMD_REQ_TASK_DETAIL = 'Request Task Detail';
  CMD_REQ_TASK_EAMIL_LIST = 'Request Task Email List';
  CMD_REQ_TASK_EAMIL_CONTENT = 'Request Task Email Content';
  CMD_EXECUTE_SAVE_TASK_DETAIL = 'Execute Save Task Detail';
  CMD_REQ_VESSEL_LIST = 'Request Vessel List';

implementation

{ TOLMsgFileRecord }

procedure TOLMsgFileRecord.Clear;
begin
  FEntryId := '';
  FStoreId := '';
  FSender := '';
  FReceiver := '';
  FCarbonCopy := '';
  FBlindCC := '';
  FSubject := '';
  FReceiveDate := 0;
  FMailItem := nil;
end;

end.
