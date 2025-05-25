unit UnitHiMTechData;

interface

uses System.Classes, UnitEnumHelper;

const
  XLS_NAME_WORKTIMETAG = '태깅';
  FILE_NAME_WORKTIMETAG = '일일근무현황';
  XLS_NAME_PAYROLLSHEET = '급여대장';
  FILE_NAME_SALARYSTATEMENT = '급여명세서';
  SHEET_NAME_PAYROLLSHEET = '금월급여대장';

type
  THiMTechDataType = (hmtdtNull, hmtdtworkTimeTag, hmtdtPayRollSheet);

const
  R_HiMTechDataType : array[Low(THiMTechDataType)..High(THiMTechDataType)] of string =
    ('', '출근 태깅 데이터', '급여 대장');

  R_HiMTechOriginalRptFileName : array[Low(THiMTechDataType)..High(THiMTechDataType)] of string =
    ('', '일일근무현황-원본.xlsx', '급여명세서-원본.xlsx');

var
  g_HiMTechDataType: TLabelledEnum<THiMTechDataType>;
  g_HiMTechOriginalRptName: TLabelledEnum<THiMTechDataType>;

implementation

initialization
  g_HiMTechDataType.InitArrayRecord(R_HiMTechDataType);
  g_HiMTechOriginalRptName.InitArrayRecord(R_HiMTechOriginalRptFileName);

end.
