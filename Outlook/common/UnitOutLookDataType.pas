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
    FSenderHandle: THandle; //메세지를 보내는 윈도우의 핸들
  end;

  TEntryIdRecord = packed record
    FEntryId,     //Email EntryId
    FStoreId,     //Email StoreId
    FEntryId4MoveRoot, //Folder EntryId
    FStoreId4MoveRoot, //Folder StoreId
    FFolderPath4Move, //Root Folder + ';' + SubFolder('\'로 구분됨)
    FFolderPath,
    FNewEntryId,
    FSubject,
    FTo,
    FHTMLBody,
    FHullNo,
    FSubFolder,
    FAttached,
    FAttachFileName: string;
    FIgnoreReceiver2pjh: Boolean; //True = 수신자가 pjh인가 비교하지 않음
    FIgnoreEmailMove2WorkFolder: Boolean; //True = Working Folder로 이동 안함
    //True = Move하고자 선택한 폴더 아래에 HullNo Folder 생성 후 생성된 폴더에 메일 이동 함
    FIsCreateHullNoFolder: Boolean;
    FSenderHandle: THandle; //메세지를 보내는 윈도우의 핸들
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

    FSenderHandle: THandle; //메세지를 보내는 윈도우의 핸들
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

    FSenderHandle: THandle; //메세지를 보내는 윈도우의 핸들
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
