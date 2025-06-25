unit UnitHiMTechExcelUtil;

interface

uses Sysutils, Dialogs, Classes, Forms, Variants,
  NxColumnClasses, NxColumns, NxGrid, NxCells,
  mormot.core.base, mormot.core.datetime, mormot.core.variants, mormot.core.data,
  mormot.core.unicode, mormot.core.text, mormot.core.os,
  UnitStringUtil, UnitFileSearchUtil, UnitExcelUtil, UnitHiMTechData;

function GetOutFileNameByDataType(ADataType: integer): string;

function ImportWorkTimeTagData2GridFromXlsFile(AFileName: string; AGrid: TNextGrid): integer;
function ImportWorkTimeTagData2GridFromString(AStr: RawByteString; AGrid: TNextGrid): integer;
function ImportPaySlipData2GridFromXlsFile(AFileName: string; AGrid: TNextGrid): integer;

function MakeHiMTechReport2ExcelByDataTypeFromGrid(AGrid: TNextGrid; ADataType: integer; ADate: TDate; ASelectedOnly: Boolean=False): integer;
function MakeDailyWorkReport2ExcelFromGrid(AOutFileName: string; AGrid: TNextGrid): integer;
function MakePaySlipReport2ExcelFromGrid(AOutFileName: string; AGrid: TNextGrid; ADate: TDate; ASelectedOnly: Boolean): integer;

function SetPayData2ExcelRptBySheetNameFromGridRow(AGrid: TNextGrid; ARow: integer; ADate: TDate; AWorkSheet: OLEVariant): integer;

implementation

uses UnitmORMotUtil2, UnitDateUtil2, UnitNextGridUtil2;

function GetOutFileNameByDataType(ADataType: integer): string;
begin
  Result := ExtractFilePath(Application.ExeName) + g_HiMTechOriginalRptName.ToString(ADataType);
end;

function ImportWorkTimeTagData2GridFromXlsFile(AFileName: string; AGrid: TNextGrid): integer;
var
  LJsonAry: string;
  LDocList: IDocList;
//  LVar: variant;
begin
  if FileExists(AFileName) then
  begin
    LJsonAry := GetJsonAryBySheetNameFromExcelFile(AFileName);
    LDocList := DocList(StringToUtf8(LJsonAry));
    NextGridFromDocList(AGrid, LDocList);
  end
  else
  begin
    ShowMessage('File [' + AFileName + ']�� �������� �ʽ��ϴ�');
    exit;
  end;
end;

function ImportWorkTimeTagData2GridFromString(AStr: RawByteString; AGrid: TNextGrid): integer;
var
  LStream: TStream;
  LTmpXlsFileName: string;
begin
  Result := -1;

  try
    LStream := RawByteStringToStream(AStr);
    //'c:\Temp\Temp.xls' �� LStream�� ������
    LTmpXlsFileName := GetFileNameFromStream(LStream);
    Result := ImportWorkTimeTagData2GridFromXlsFile(LTmpXlsFileName, AGrid);
  finally
    LStream.Free
  end;
end;

function ImportPaySlipData2GridFromXlsFile(AFileName: string; AGrid: TNextGrid): integer;
var
  LJsonAry: string;
  LDocList: IDocList;
begin
  if FileExists(AFileName) then
  begin
    LJsonAry := GetJsonAryBySheetNameFromExcelFile(AFileName, SHEET_NAME_PAYROLLSHEET, 'B', '5', 'B', '7', True);
    LDocList := DocList(StringToUtf8(LJsonAry));
    NextGridFromDocList(AGrid, LDocList);
  end
  else
  begin
    ShowMessage('File [' + AFileName + ']�� �������� �ʽ��ϴ�');
    exit;
  end;
end;

function MakeHiMTechReport2ExcelByDataTypeFromGrid(AGrid: TNextGrid; ADataType: integer;
  ADate: TDate; ASelectedOnly: Boolean): integer;
var
  LOutFileName, LTempFileName: string;
  LFileCopySuccess: Boolean;
  LRow: integer;
begin
  LOutFileName := GetOutFileNameByDataType(ADataType);
  EnsureDirectoryExists('c:\temp\');
  LTempFileName := 'c:\temp\' + ChangeFileExt(ExtractFileName(LOutFileName), '-' + FormatDateTime('yyyymmddhhnnss' , now) + '.xlsx');
  LTempFileName := StringReplace(LTempFileName, '-����', '', [rfReplaceAll]);
  LFileCopySuccess := CopyFile(LOutFileName, LTempFileName, False);

  if LFileCopySuccess then
  begin
    case g_HiMTechDataType.ToType(ADataType) of
      hmtdtworkTimeTag: begin
        MakeDailyWorkReport2ExcelFromGrid(LTempFileName, AGrid);
      end;
      hmtdtPayRollSheet: begin
        MakePaySlipReport2ExcelFromGrid(LTempFileName, AGrid, ADate, ASelectedOnly);
      end;
    end; //case
  end
  else
    ShowMessage('���� ���� ���� : [' + LOutFileName + '] --> [' + LTempFileName + ']');
end;

function MakeDailyWorkReport2ExcelFromGrid(AOutFileName: string; AGrid: TNextGrid): integer;
var
  LRange: OleVariant;
  LWorksheet: OleVariant;

  i, LRow, LCol, LFindFromRow: integer;
  LEmployeeName, LWorkBegin, LWorkEnd, LWorkPeriod, LWorkOT, LAttendance,
  LRangeStr: string;
  LVarAry: variant;
begin
  Result := -1;

  LWorksheet := GetWorkSheetByNameFromExcelFile(AOutFileName, '');

  if VarIsNull(LWorksheet) then
  begin
    ShowMessage('Worksheet�� �������� �ʽ��ϴ�');
    exit;
  end;

  try
    LRow := GetLastRowNumFromExcel(LWorksheet);

    for i := 7 to LRow do
    begin
      LRange := LWorksheet.range['F'+IntToStr(i)];
      LEmployeeName := LRange.FormulaR1C1;

      if LEmployeeName = '' then
        Break;
    end;

    if i >= LRow then
      exit;

    LRangeStr := 'F7:M' + IntToStr(i-1);
    LVarAry := ReadExcelRangeToVarArrayFromWorkSheet(LWorksheet, LRangeStr);

    for LRow := VarArrayLowBound(LVarAry, 1) to VarArrayHighBound(LVarAry, 1) do
    begin
      LEmployeeName := VarToStr(LVarAry[LRow, 1]);

      LCol := GetColIdxByColCaptionFromGrid(AGrid, '����');
      LFindFromRow := 0;
      LFindFromRow := GetRowIndexFromFindNext(AGrid, LEmployeeName, LCol, LFindFromRow);

      if LFindFromRow = -1 then
        Continue;

      LWorkBegin := GetCellDataByColCaptionFromGrid(AGrid, '�ٹ� ����', LFindFromRow);
      LWorkEnd := GetCellDataByColCaptionFromGrid(AGrid, '�ٹ� ����', LFindFromRow);
      LWorkPeriod := GetCellDataByColCaptionFromGrid(AGrid, '�ٷ� �ð�', LFindFromRow);
      LWorkOT := GetCellDataByColCaptionFromGrid(AGrid, '�ʰ� �ٹ�', LFindFromRow);
      LAttendance := GetCellDataByColCaptionFromGrid(AGrid, '����', LFindFromRow);

      LVarAry[LRow, 4] := LAttendance;
//      LVarAry[LRow, 3] := '';
      LVarAry[LRow, 6] := LWorkBegin;
      LVarAry[LRow, 7] := LWorkEnd;
      LVarAry[LRow, 8] := AddTimeStrings(LWorkPeriod, LWorkOT);

//      for LCol := VarArrayLowBound(LVarAry, 2) to VarArrayHighBound(LVarAry, 2) do
//      begin
//        VarToStr(LVarAry[LRow, LCol]);
//      end;
    end;//for

    LWorksheet.Range[LRangeStr].Value2 := LVarAry;
    ShowMessage('���Ͼ�����Ȳ ���� �Ϸ�');
  finally
//    LExcel.WorkBooks.Close;
//    LExcel.quit;
//    LExcel:=unassigned;
  end;
end;

function MakePaySlipReport2ExcelFromGrid(AOutFileName: string; AGrid: TNextGrid;
  ADate: TDate;  ASelectedOnly: Boolean): integer;
var
  LWorksheet: OleVariant;
  LExcel: OleVariant;
  LWorkBook: OleVariant;

  i: integer;
  LEmployeeName: string;
begin
  Result := -1;

  if not FileExists(AOutFileName) then
  begin
    ShowMessage('File(' + AOutFileName + ')�� �������� �ʽ��ϴ�');
    exit;
  end;

  LExcel := GetActiveExcelOleObject(True);
  LWorkBook := LExcel.Workbooks.Open(AOutFileName);
  LExcel.Visible := true;
//  LWorksheet := LExcel.ActiveSheet;

  for i := 0 to AGrid.RowCount - 1 do
  begin
    if ASelectedOnly then
      if not AGrid.Row[i].Selected then
        Continue;

    LEmployeeName := GetCellDataByColCaptionFromGrid(AGrid, '����', i);
    LWorksheet := CopySheet2WorkBookByName(LWorkBook, '1', LEmployeeName);

    if VarIsNull(LWorksheet) then
    begin
      Continue;
    end;

    SetPayData2ExcelRptBySheetNameFromGridRow(AGrid, i, ADate, LWorkSheet);
  end;//for

//  try
//      LWorkBegin := GetCellDataByColCaptionFromGrid(AGrid, '�ٹ� ����', LFindFromRow);
//      LWorkEnd := GetCellDataByColCaptionFromGrid(AGrid, '�ٹ� ����', LFindFromRow);
//      LWorkPeriod := GetCellDataByColCaptionFromGrid(AGrid, '�ٷ� �ð�', LFindFromRow);
//      LWorkOT := GetCellDataByColCaptionFromGrid(AGrid, '�ʰ� �ٹ�', LFindFromRow);
//      LAttendance := GetCellDataByColCaptionFromGrid(AGrid, '����', LFindFromRow);
//
//    ShowMessage('�޿���ǥ ���� �Ϸ�');
//  finally
////    LExcel.WorkBooks.Close;
////    LExcel.quit;
////    LExcel:=unassigned;
//  end;
end;

function SetPayData2ExcelRptBySheetNameFromGridRow(AGrid: TNextGrid; ARow: integer;
  ADate: TDate; AWorkSheet: OLEVariant): integer;
var
  LRange, LRange2: OleVariant;
  i, ItemCount, LCol, LCol2, LRangeRow: integer;
  LEmployeeName, LEntryDate, LWagePerHour, LPaidDayOff, LAnnualLeave, LWeeklyLeave,
  LOverTime, LWorkHour, LWorkHour_Night, LWorkHour_Holiday, LWorkHour_Add,
  LRangeStr, LDateStr, LValue, LRangeStr2: string;
begin
  LDateStr := FormatDateTime('yyyy�� mm�� �޿���ǥ', ADate);

  LEmployeeName := GetCellDataByColCaptionFromGrid(AGrid, '����', ARow);
  LEntryDate := GetCellDataByColCaptionFromGrid(AGrid, '�Ի�����', ARow);
  LWagePerHour := GetCellDataByColCaptionFromGrid(AGrid, '�ñ�', ARow);
  LPaidDayOff := GetCellDataByColCaptionFromGrid(AGrid, '��������', ARow);
  LAnnualLeave := GetCellDataByColCaptionFromGrid(AGrid, '����', ARow);
  LWeeklyLeave := GetCellDataByColCaptionFromGrid(AGrid, '����', ARow);
  LOverTime := GetCellDataByColCaptionFromGrid(AGrid, '����ð�', ARow);
  LWorkHour := GetCellDataByColCaptionFromGrid(AGrid, '�ٷνð�', ARow);
  LWorkHour_Night := GetCellDataByColCaptionFromGrid(AGrid, '�߰��ٷνð�', ARow);
  LWorkHour_Holiday := GetCellDataByColCaptionFromGrid(AGrid, '���ϱٷνð�', ARow);
  LWorkHour_Add := GetCellDataByColCaptionFromGrid(AGrid, '���޽ð�', ARow);

  LRange := AWorkSheet.range['D2'];
  LRange.FormulaR1C1 := LDateStr;
  LRange := AWorkSheet.range['F5'];
  LRange.FormulaR1C1 := LEmployeeName;
  LRange := AWorkSheet.range['O5'];
  LRange.FormulaR1C1 := LEntryDate;
  LRange := AWorkSheet.range['E7'];
  LRange.FormulaR1C1 := LWagePerHour;
  LRange := AWorkSheet.range['A9'];
  LRange.FormulaR1C1 := LWorkHour;
  LRange := AWorkSheet.range['C9'];
  LRange.FormulaR1C1 := LWorkHour_Night;
  LRange := AWorkSheet.range['E9'];
  LRange.FormulaR1C1 := LWorkHour_Holiday;
  LRange := AWorkSheet.range['G9'];
  LRange.FormulaR1C1 := LOverTime;
  LRange := AWorkSheet.range['I9'];
  LRange.FormulaR1C1 := LWorkHour_Add;
  LRange := AWorkSheet.range['K9'];
  LRange.FormulaR1C1 := LPaidDayOff;
  LRange := AWorkSheet.range['M9'];
  LRange.FormulaR1C1 := LAnnualLeave;
  LRange := AWorkSheet.range['O9'];
  LRange.FormulaR1C1 := LWeeklyLeave;

  //���� �׸�(��������) ä���
  LCol := GetColIdxByColCaptionFromGrid(AGrid, '�ð��Ѿ�');

  if LCol > 0 then
  begin
    //Header Caption�� '�ð��Ѿ�' ���� '�����հ�' ���� ���� �׸���
    Inc(LCol);

    //���� �׸��� ������ Column Index ������
    LCol2 := GetColIdxByColCaptionFromGrid(AGrid, '�����հ�');

    //Header Caption '�����հ�' �� '�ð��Ѿ�' �ڿ� �־�� ��
    if (LCol2 > 0) and (LCol < LCol2) then
    begin
      Dec(LCol2);
      ItemCount := LCol2 - LCol + 1;

      //���� �׸��� 10�� �̻��̸� ���� Row �߰�
      if ItemCount > 10 then
      begin
        for i := 10 to ItemCount do
          XlsRangeCopyNInsert2WS(AWorkSheet, 'B19:P19', 'B20:P20');
      end;

      LRangeRow := 11;

      for i := LCol to LCol2 do
      begin
        //���� �׸� ����
        LRangeStr := 'L' + IntToStr(LRangeRow);
        LRange := AWorkSheet.range[LRangeStr];
        LRange.FormulaR1C1 := AGrid.Columns.Item[i].Header.Caption;

        //���� �ݾ� ����
        LRangeStr := 'P' + IntToStr(LRangeRow);
        LRange := AWorkSheet.range[LRangeStr];
        LRange.FormulaR1C1 := AGrid.Cells[i, ARow];

        Inc(LRangeRow);
      end;
    end;
  end;

  //���� �׸� ä���
  LCol := GetColIdxByColCaptionFromGrid(AGrid, '�ѱ޿���');

  if LCol > 0 then
  begin
    //Header Caption�� '�ѱ޿���' ���� Grid�� ������ Column ���� ���� �׸���
    Inc(LCol);
    //25����� �����ؼ� '���ο���'�� �ִ� ���� �˻���
    LRangeRow := GetRowidxByCellValueFromWS(AWorkSheet, '���ο���', 25);
    LRangeStr2 := 'A' + IntToStr(LRangeRow + 4) + ':T' + IntToStr(LRangeRow + 4);

    LCol2 := 0;

    for i := LCol to AGrid.Columns.Count - 1 do
    begin
      LDateStr := AGrid.Columns.Item[i].Header.Caption;
      LValue := AGrid.Cells[i, ARow];

      if (LDateStr = '���ο���') or (LDateStr = '�ǰ�����') or
        (LDateStr = '��뺸��') or (LDateStr = '���ټ�') or (LDateStr = '�ֹμ�') then
        LRangeStr := 'F';

      if LRangeRow <> -1 then
      begin
        if LDateStr = '���ο���' then
          ItemCount := LRangeRow
        else if LDateStr = '�ǰ�����' then
          ItemCount := LRangeRow+1
        else if LDateStr = '��뺸��' then
          ItemCount := LRangeRow+2
        else if LDateStr = '���ټ�' then
          ItemCount := LRangeRow+3
        else if LDateStr = '�ֹμ�' then
          ItemCount := LRangeRow+4
        else
        begin
          if LValue = '' then
            Continue;

          if LDateStr = '�����װ�' then
            Break;

          ItemCount := LRangeRow + LCol2;

          //�������� �� ���������� �⺻�� 5����, ��� 5���� ũ�� ���� �߰� �ؾ���
          if LCol2 > 4 then
          begin
            LRangeStr := 'A' + IntToStr(LRangeRow + LCol2-1) + ':T' + IntToStr(LRangeRow + LCol2-1);
            LRangeStr2 := 'A' + IntToStr(LRangeRow + LCol2) + ':T' + IntToStr(LRangeRow + LCol2);
            XlsRangeCopyNInsert2WS(AWorkSheet, LRangeStr, LRangeStr2);
            LRange := AWorkSheet.range[LRangeStr2];
            LRange.FormulaR1C1 := '';  //���� �߰��� ���� ������ ���
          end;

          //���� �׸��� ������ �߰���
          LRangeStr := 'K' + IntToStr(ItemCount);
          LRange := AWorkSheet.range[LRangeStr];
          LRange.FormulaR1C1 := LDateStr;

          LRangeStr := 'P';

          Inc(LCol2);
        end;

        LRangeStr := LRangeStr + IntToStr(ItemCount);
        LRange := AWorkSheet.range[LRangeStr];
        LRange.FormulaR1C1 := LValue;
      end;
    end;//for

    LRange := AWorkSheet.range[LRangeStr2];
    //xlEdgeBottom, xlContinuous, xlThick
    SetExcelCellRangeSingleBorder(LRange, 9, 1, 4);
  end;
end;

end.
