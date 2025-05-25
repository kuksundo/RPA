unit UnitHiMTechData;

interface

uses System.Classes, UnitEnumHelper;

const
  XLS_NAME_WORKTIMETAG = '�±�';
  FILE_NAME_WORKTIMETAG = '���ϱٹ���Ȳ';
  XLS_NAME_PAYROLLSHEET = '�޿�����';
  FILE_NAME_SALARYSTATEMENT = '�޿�����';
  SHEET_NAME_PAYROLLSHEET = '�ݿ��޿�����';

type
  THiMTechDataType = (hmtdtNull, hmtdtworkTimeTag, hmtdtPayRollSheet);

const
  R_HiMTechDataType : array[Low(THiMTechDataType)..High(THiMTechDataType)] of string =
    ('', '��� �±� ������', '�޿� ����');

  R_HiMTechOriginalRptFileName : array[Low(THiMTechDataType)..High(THiMTechDataType)] of string =
    ('', '���ϱٹ���Ȳ-����.xlsx', '�޿�����-����.xlsx');

var
  g_HiMTechDataType: TLabelledEnum<THiMTechDataType>;
  g_HiMTechOriginalRptName: TLabelledEnum<THiMTechDataType>;

implementation

initialization
  g_HiMTechDataType.InitArrayRecord(R_HiMTechDataType);
  g_HiMTechOriginalRptName.InitArrayRecord(R_HiMTechOriginalRptFileName);

end.
