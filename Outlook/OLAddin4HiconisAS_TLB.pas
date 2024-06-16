unit OLAddin4HiconisAS_TLB;

{$TYPEDADDRESS OFF}

interface

uses SysUtils, ComObj, ComServ, ActiveX, Variants;

const
  OLAddin4HiconisASMajorVersion = 1;
  OLAddin4HiconisASMinorVersion = 0;

  LIBID_OLAddin4HiconisAS: TGUID = '{42B5D5E9-6A47-4911-822A-58BD6F9BE26D}';

  IID_ICoOLAddin4HiconisAS: TGUID = '{4A93F7BE-E3D1-49B3-97ED-2064896668C7}';
  CLASS_CoOLAddin4HiconisAS: TGUID = '{E2D330BF-D398-4D0A-AAC1-F64AFE47A835}';

type
  ICoOLAddin4HiconisAS = interface(IDispatch)
    ['{4A93F7BE-E3D1-49B3-97ED-2064896668C7}']
  end;

  ICoOLAddin4HiconisASDisp = dispinterface
    ['{4A93F7BE-E3D1-49B3-97ED-2064896668C7}']
  end;

implementation

end.
