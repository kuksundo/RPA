unit UnitOutLookDataType;

interface

uses
  mormot.core.base,
  mormot.orm.base,
  UnitEnumHelper;

type
  TOLRespondRec = packed record
    FID: integer;
    FMsg: string;
  end;

  TOLCommandKind = (olckInitVar, olckAddAppointment, olckMoveMail2Folder, olckGetFolderList);
  TOLRespondKind = (olrkMAPIFolderList, olrkLog);

const
  MEMO_LOG_MAX_LINE_COUNT = 100;

  R_OLCommandKind : array[Low(TOLCommandKind)..High(TOLCommandKind)] of string =
    (
      'Init Var',
      'Add Appointment',
      'Move Mail To Folder',
      'Get FolderList'
    );
  R_OLRespondKind : array[Low(TOLRespondKind)..High(TOLRespondKind)] of string =
    ('MAPIFolder List', 'Log');

var
  g_OLCommandKind: TLabelledEnum<TOLCommandKind>;
  g_OLRespondKind: TLabelledEnum<TOLRespondKind>;

implementation

initialization
  g_OLCommandKind.InitArrayRecord(R_OLCommandKind);
  g_OLRespondKind.InitArrayRecord(R_OLRespondKind);

end.
