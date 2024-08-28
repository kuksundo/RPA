unit UnitOutLookDataType;

interface

uses Outlook_TLB,
  mormot.core.base,
  mormot.orm.base,
  UnitEnumHelper;

type
  TOLRespondRec = packed record
    FID: integer;
    FMsg: string;
    FSenderHandle: THandle; //�޼����� ������ �������� �ڵ�
  end;

  TEntryIdRecord = packed record
    FEntryId,     //Email EntryId
    FStoreId,     //Email StoreId
    FEntryId4MoveRoot, //Folder EntryId
    FStoreId4MoveRoot, //Folder StoreId
    FFolderPath4Move, //Root Folder + ';' + SubFolder('\'�� ���е�)
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
    FSenderHandle: THandle; //�޼����� ������ �������� �ڵ�
//    FIsShowMailContents: Boolean; //True = Mail Display
  end;

  TOLMsgFileRecord = packed record
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

  TOLMailRec = packed record
    Body: WideString;
    Categories: WideString;
    Companies: WideString;
    CreationTime: TDateTime;
    ReceivedTime: TDateTime;
    EntryID: WideString;
//    Size: Integer;
    Subject: WideString;
    BCC: WideString;
    CC: WideString;
    HTMLBody: WideString;
    Recipients: WideString;
    SenderName: WideString;
    BodyFormat: LongWord; //OlBodyFormat
    SenderEmailAddress: WideString;
    To_: WideString;

    FSenderHandle: THandle; //�޼����� ������ �������� �ڵ�
  end;

  TOLObjectRec = packed record
    OLObjectKind: LongWord; //TOLObjectKind

    EntryID: WideString;
    Body: WideString;
    Categories: WideString;
    Companies: WideString;
    ConversationIndex: WideString;
    ConversationTopic: WideString;
    CreationTime: TDateTime;
    LastModificationTime: TDateTime;
    Mileage: WideString;
    Subject: WideString;
    Duration: Integer;
    Start: TDateTime;
    End_: TDateTime;
    IsOnlineMeeting: WordBool;
    Location: WideString;
    OptionalAttendees: WideString;
    Organizer: WideString;
    ReminderSet: WordBool;
    ReminderMinutesBeforeStart: Integer;
    RequiredAttendees: WideString;

    FSenderHandle: THandle; //�޼����� ������ �������� �ڵ�
  end;

  TOLObjectKind = (olobjAppointment, olobjTask, olobjMeeting, olobjEvent, olobjNote, olobjContact, olobjVCard);

  TOLCommandKind = (
    olckInitVar,
    olckAddObject,
    olckMoveMail2Folder,
    olckGetFolderList,
    olckGetSelectedMailItemFromExplorer,
    olckShowMailContents,
    olckShowObject,
    olckCreateMail,
    olcGotoFolder,
    olckFinal);
  TOLRespondKind = (
    olrkInitVar,
    olrkAddObject,
    olrkMAPIFolderList,
    olrkLog,
    olrkSelMail4Explore,
    olrkMoveMail2Folder,
    olrkShowObject,
    olrkCreateMail,
    olrkGotoFolder,
    olrkFinal
    );

const
  MEMO_LOG_MAX_LINE_COUNT = 100;

  R_OLCommandKind : array[Low(TOLCommandKind)..High(TOLCommandKind)] of string =
    (
      'Init Var',
      'Add Object',
      'Move Mail To Folder',
      'Get FolderList',
      'Get Selected MailItem From Explorer',
      'Show Mail Contents',
      'Show Object',
      'Create Mail',
      'Go To Folder',
      ''
    );
  R_OLRespondKind : array[Low(TOLRespondKind)..High(TOLRespondKind)] of string =
    (
      'Init Outlook OK',
      'Object Added',
      'MAPIFolder List',
      'Log',
      'Selected Mail Item',
      'Move Mail to Folder',
      'Show Object',
      'Create Mail',
      'Go To Folder',
      ''
    );

function GetOLObjItemFromOLKind(const AOLObjKind: integer): LongWord;

var
  g_OLCommandKind: TLabelledEnum<TOLCommandKind>;
  g_OLRespondKind: TLabelledEnum<TOLRespondKind>;

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

function GetOLObjItemFromOLKind(const AOLObjKind: integer): LongWord;
begin
  case TOLObjectKind(AOLObjKind) of
    olobjAppointment: Result := olAppointmentItem;
    olobjTask: Result := olTaskItem;
    olobjMeeting: Result := olAppointmentItem;
    olobjEvent: Result := olAppointmentItem;
    olobjNote: Result := olNoteItem;
    olobjContact: Result := olContactItem;
    olobjVCard: Result := olAppointmentItem;
  end;
end;

{ TEntryIdRecord }

initialization
  g_OLCommandKind.InitArrayRecord(R_OLCommandKind);
  g_OLRespondKind.InitArrayRecord(R_OLRespondKind);

end.
