using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelRowSplitter
{
    public partial class Settlement : Form
    {
        private string outputDirectory;
        private string[] headerRow;

        public Settlement()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 파일 첨부 버튼 클릭 이벤트 핸들러
        /// </summary>
        private void btnAttachFile_Click(object sender, EventArgs e)
        {
            string filePath = SelectExcelFile();
            if (string.IsNullOrEmpty(filePath)) return;

            string outputFolder = SelectOutputFolder();
            if (string.IsNullOrEmpty(outputFolder)) return;

            outputDirectory = outputFolder;
            ProcessExcelFile(filePath);
        }

        /// <summary>
        /// Excel 파일 선택 대화상자 표시
        /// </summary>
        /// <returns>선택된 파일 경로 또는 null</returns>
        private string SelectExcelFile()
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xls;*.xlsx",
                Title = "전체 데이터 엑셀파일 선택하세요"
            })
            {
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    lblFilePath.Text = openFileDialog.FileName;
                    return openFileDialog.FileName;
                }
                return null;
            }
        }

        /// <summary>
        /// 출력 폴더 선택 대화상자 표시
        /// </summary>
        /// <returns>선택된 폴더 경로 또는 null</returns>
        private string SelectOutputFolder()
        {
            using (FolderBrowserDialog folderDialog = new FolderBrowserDialog
            {
                Description = "정산서를 배포할 폴더 선택하세요"
            })
            {
                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    return folderDialog.SelectedPath;
                }
                MessageBox.Show("폴더가 선택되지 않았습니다. 작업이 취소되었습니다.");
                return null;
            }
        }

        /// <summary>
        /// Excel 파일 처리 메서드
        /// </summary>
        /// <param name="filePath">처리할 Excel 파일 경로</param>
        private void ProcessExcelFile(string filePath)
        {
            try
            {
                using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    IWorkbook workbook = new XSSFWorkbook(fs);
                    if (!ValidateWorkbook(workbook)) return;

                    ISheet[] sheets = LoadSheets(workbook);
                    ProcessSheets(sheets);
                }

                MessageBox.Show("정산서 데이터추출 완료!");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"오류가 발생했습니다: {ex.Message}");
            }
        }

        /// <summary>
        /// Workbook 유효성 검사
        /// </summary>
        /// <param name="workbook">검사할 Workbook</param>
        /// <returns>유효성 여부</returns>
        private bool ValidateWorkbook(IWorkbook workbook)
        {
            if (workbook.NumberOfSheets < 6)
            {
                MessageBox.Show("엑셀 파일에 시트가 6개 이상 존재하지 않습니다.");
                return false;
            }
            return true;
        }

        /// <summary>
        /// Workbook에서 필요한 시트들을 로드
        /// </summary>
        /// <param name="workbook">시트를 로드할 Workbook</param>
        /// <returns>로드된 시트 배열</returns>
        private ISheet[] LoadSheets(IWorkbook workbook)
        {
            return new ISheet[]
            {
                workbook.GetSheetAt(0),
                workbook.GetSheetAt(1),
                workbook.GetSheetAt(2),
                workbook.GetSheetAt(3),
                workbook.GetSheetAt(4),
                workbook.GetSheetAt(5)
            };
        }

        /// <summary>
        /// 시트 처리 메서드
        /// </summary>
        /// <param name="sheets">처리할 시트 배열</param>
        private void ProcessSheets(ISheet[] sheets)
        {
            ISheet sheet1 = sheets[0];
            int rowCount = sheet1.PhysicalNumberOfRows;

            headerRow = ExtractHeaderRow(sheet1);
            SetupProgressBar(rowCount);

            for (int row = 1; row < rowCount; row++)
            {
                ProcessRow(row, sheets);
            }
        }

        /// <summary>
        /// 헤더 행 추출
        /// </summary>
        /// <param name="sheet">헤더를 추출할 시트</param>
        /// <returns>추출된 헤더 배열</returns>
        private string[] ExtractHeaderRow(ISheet sheet)
        {
            IRow header = sheet.GetRow(0);
            string[] headerData = new string[32];
            for (int col = 0; col < 32; col++)
            {
                headerData[col] = header.GetCell(col)?.ToString();
            }
            return headerData;
        }

        /// <summary>
        /// 진행 상황 표시바 설정
        /// </summary>
        /// <param name="maxValue">최대 값</param>
        private void SetupProgressBar(int maxValue)
        {
            progressBar.Maximum = maxValue - 1;
            progressBar.Value = 0;
        }

        /// <summary>
        /// 개별 행 처리
        /// </summary>
        /// <param name="rowIndex">처리할 행 인덱스</param>
        /// <param name="sheets">시트 배열</param>
        private void ProcessRow(int rowIndex, ISheet[] sheets)
        {
            try
            {
                string[] rowData = ExtractRowData(sheets[0].GetRow(rowIndex));
                if (!ValidateRowData(rowData)) return;

                string fileName = GenerateFileName(rowData);
                SaveRowToNewExcelFile(rowData, fileName, sheets);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"행 {rowIndex} 처리 중 오류 발생: {ex.Message}");
            }
            finally
            {
                progressBar.Value++;
            }
        }

        /// <summary>
        /// 행 데이터 추출
        /// </summary>
        /// <param name="row">데이터를 추출할 행</param>
        /// <returns>추출된 행 데이터 배열</returns>
        private string[] ExtractRowData(IRow row)
        {
            string[] rowData = new string[32];
            for (int col = 0; col < 32; col++)
            {
                rowData[col] = row.GetCell(col)?.ToString();
            }
            return rowData;
        }

        /// <summary>
        /// 행 데이터 유효성 검사
        /// </summary>
        /// <param name="rowData">검사할 행 데이터</param>
        /// <returns>유효성 여부</returns>
        private bool ValidateRowData(string[] rowData)
        {
            return !string.IsNullOrWhiteSpace(rowData[0]) &&
                   !string.IsNullOrWhiteSpace(rowData[1]) &&
                   !string.IsNullOrWhiteSpace(rowData[2]) &&
                   !string.IsNullOrWhiteSpace(rowData[3]) &&
                   !string.IsNullOrWhiteSpace(rowData[5]);
        }

        /// <summary>
        /// 파일명 생성
        /// </summary>
        /// <param name="rowData">파일명 생성에 사용할 행 데이터</param>
        /// <returns>생성된 파일명</returns>
        private string GenerateFileName(string[] rowData)
        {
            string fileName = $"{rowData[0]}~{rowData[1]}_{rowData[5]}_{rowData[3]}_{rowData[2]}.xlsx";
            return CleanFileName(fileName);
        }

        /// <summary>
        /// 파일명에서 특수문자 제거
        /// </summary>
        /// <param name="fileName">정리할 파일명</param>
        /// <returns>특수문자가 제거된 파일명</returns>
        private string CleanFileName(string fileName)
        {
            return Regex.Replace(fileName, @"[\\/:*?""<>|]", string.Empty);
        }

        /// <summary>
        /// 새로운 엑셀 파일로 데이터 저장
        /// </summary>
        /// <param name="rowData">저장할 행 데이터</param>
        /// <param name="fileName">생성할 파일명</param>
        /// <param name="sourceSheets">원본 시트 배열</param>
        private void SaveRowToNewExcelFile(string[] rowData, string fileName, ISheet[] sourceSheets)
        {
            IWorkbook newWorkbook = new XSSFWorkbook();
            ISheet[] newSheets = CreateNewSheets(newWorkbook);

            string sheet1C2Value = rowData[2];

            AddHeaderAndDataToSheet1(newSheets[0], rowData);
            CopyAndFilterOtherSheets(sourceSheets, newSheets, sheet1C2Value);

            RemoveSpecificColumns(newSheets);
            ProcessRanges(newSheets);
            RemoveEmptyRowsFromAllSheets(newSheets);

            SaveWorkbook(newWorkbook, fileName);
        }

        /// <summary>
        /// 새 워크북에 시트 생성
        /// </summary>
        /// <param name="workbook">새 워크북</param>
        /// <returns>생성된 시트 배열</returns>
        private ISheet[] CreateNewSheets(IWorkbook workbook)
        {
            return new ISheet[]
            {
                workbook.CreateSheet("Sheet1"),
                workbook.CreateSheet("Sheet2"),
                workbook.CreateSheet("Sheet3"),
                workbook.CreateSheet("Sheet4"),
                workbook.CreateSheet("Sheet5"),
                workbook.CreateSheet("Sheet6")
            };
        }

        /// <summary>
        /// 시트1에 헤더와 데이터 추가
        /// </summary>
        /// <param name="sheet">대상 시트</param>
        /// <param name="rowData">추가할 행 데이터</param>
        private void AddHeaderAndDataToSheet1(ISheet sheet, string[] rowData)
        {
            // 헤더 행 추가
            IRow headerRowInNewFile = sheet.CreateRow(0);
            for (int col = 0; col < headerRow.Length; col++)
            {
                headerRowInNewFile.CreateCell(col).SetCellValue(headerRow[col]);
            }

            // 데이터 행 추가
            IRow newRow = sheet.CreateRow(1);
            for (int col = 0; col < rowData.Length; col++)
            {
                newRow.CreateCell(col).SetCellValue(rowData[col]);
            }
        }

        /// <summary>
        /// 다른 시트들을 복사하고 필터링
        /// </summary>
        /// <param name="sourceSheets">원본 시트 배열</param>
        /// <param name="targetSheets">대상 시트 배열</param>
        /// <param name="filterValue">필터링 기준 값</param>
        private void CopyAndFilterOtherSheets(ISheet[] sourceSheets, ISheet[] targetSheets, string filterValue)
        {
            CopyAndFilterSheet(sourceSheets[1], targetSheets[1], filterValue, 2);
            CopyAndFilterSheet(sourceSheets[2], targetSheets[2], filterValue, 3);
            CopyAndFilterSheet(sourceSheets[3], targetSheets[3], filterValue, 3);
            CopyAndFilterSheet(sourceSheets[4], targetSheets[4], filterValue, 3);
            CopyAndFilterSheet(sourceSheets[4], targetSheets[5], filterValue, 0);
        }

        /// <summary>
        /// 시트 복사 및 필터링
        /// </summary>
        /// <param name="sourceSheet">원본 시트</param>
        /// <param name="targetSheet">대상 시트</param>
        /// <param name="compareValue">비교 값</param>
        /// <param name="compareColumnIndex">비교할 열 인덱스</param>
        private void CopyAndFilterSheet(ISheet sourceSheet, ISheet targetSheet, string compareValue, int compareColumnIndex)
        {
            CopyHeader(sourceSheet, targetSheet);
            CopyFilteredData(sourceSheet, targetSheet, compareValue, compareColumnIndex);
        }

        /// <summary>
        /// 헤더 행 복사
        /// </summary>
        /// <param name="sourceSheet">원본 시트</param>
        /// <param name="targetSheet">대상 시트</param>
        private void CopyHeader(ISheet sourceSheet, ISheet targetSheet)
        {
            IRow headerRow = sourceSheet.GetRow(0);
            IRow newHeaderRow = targetSheet.CreateRow(0);

            for (int col = 0; col < headerRow.LastCellNum; col++)
            {
                newHeaderRow.CreateCell(col).SetCellValue(headerRow.GetCell(col).ToString());
            }
        }

        /// <summary>
        /// 필터링된 데이터 복사
        /// </summary>
        /// <param name="sourceSheet">원본 시트</param>
        /// <param name="targetSheet">대상 시트</param>
        /// <param name="compareValue">비교 값</param>
        /// <param name="compareColumnIndex">비교할 열 인덱스</param>
        private void CopyFilteredData(ISheet sourceSheet, ISheet targetSheet, string compareValue, int compareColumnIndex)
        {
            int targetRowIndex = 1;

            for (int i = 1; i <= sourceSheet.LastRowNum; i++)
            {
                IRow sourceRow = sourceSheet.GetRow(i);
                if (sourceRow == null) continue;

                ICell compareCell = sourceRow.GetCell(compareColumnIndex);
                if (compareCell != null && compareCell.ToString() == compareValue)
                {
                    CopyRow(sourceRow, targetSheet.CreateRow(targetRowIndex++));
                }
            }
        }

        /// <summary>
        /// 행 복사
        /// </summary>
        /// <param name="sourceRow">원본 행</param>
        /// <param name="targetRow">대상 행</param>
        private void CopyRow(IRow sourceRow, IRow targetRow)
        {
            for (int j = 0; j < sourceRow.LastCellNum; j++)
            {
                ICell sourceCell = sourceRow.GetCell(j);
                ICell targetCell = targetRow.CreateCell(j);
                if (sourceCell != null)
                {
                    targetCell.SetCellValue(sourceCell.ToString());
                }
            }
        }

        /// <summary>
        /// 특정 열 제거
        /// </summary>
        /// <param name="sheets">처리할 시트 배열</param>
        private void RemoveSpecificColumns(ISheet[] sheets)
        {
            RemoveColumns(sheets[0], new[] { 6, 7, 8, 10, 11 });
            RemoveColumns(sheets[1], new[] { 9, 10, 12, 13 });
        }

        /// <summary>
        /// 특정 열 제거
        /// </summary>
        /// <param name="sheet">처리할 시트</param>
        /// <param name="columnIndexes">제거할 열 인덱스 배열</param>
        private void RemoveColumns(ISheet sheet, int[] columnIndexes)
        {
            foreach (var columnIndex in columnIndexes.OrderByDescending(c => c))
            {
                foreach (IRow row in sheet)
                {
                    if (row.GetCell(columnIndex) != null)
                    {
                        row.RemoveCell(row.GetCell(columnIndex));
                    }
                }
            }
        }

        /// <summary>
        /// 모든 시트의 특정 범위 처리
        /// </summary>
        /// <param name="sheets">처리할 시트 배열</param>
        private void ProcessRanges(ISheet[] sheets)
        {
            ProcessRange(sheets[0], "G2", "AD");
            ProcessRange(sheets[1], "I2", "AD");
            ProcessRange(sheets[2], "E2", "N");
            ProcessRange(sheets[3], "G2", "G");
            ProcessRange(sheets[4], "G2", "G");
            ProcessRange(sheets[5], "G2", "O");
            ProcessRange(sheets[5], "T2", "Z");
        }

        /// <summary>
        /// 특정 범위의 셀 처리 (숫자 형식 변경)
        /// </summary>
        /// <param name="sheet">처리할 시트</param>
        /// <param name="startCellAddress">시작 셀 주소</param>
        /// <param name="endColumnLetter">끝 열 문자</param>
        private void ProcessRange(ISheet sheet, string startCellAddress, string endColumnLetter)
        {
            int startRow = CellReference.ConvertCellReference(startCellAddress).Row;
            int startColumn = CellReference.ConvertCellReference(startCellAddress).Col;
            int lastRow = sheet.LastRowNum;
            int endColumn = CellReference.ConvertCellReference(endColumnLetter + "1").Col;

            for (int row = startRow; row <= lastRow; row++)
            {
                IRow currentRow = sheet.GetRow(row);
                if (currentRow == null) continue;

                for (int col = startColumn; col <= endColumn; col++)
                {
                    ProcessCell(sheet, currentRow.GetCell(col));
                }
            }
        }

        /// <summary>
        /// 개별 셀 처리 (숫자 형식 변경)
        /// </summary>
        /// <param name="sheet">처리할 시트</param>
        /// <param name="cell">처리할 셀</param>
        private void ProcessCell(ISheet sheet, ICell cell)
        {
            if (cell != null && cell.CellType == CellType.String && double.TryParse(cell.StringCellValue, out double result))
            {
                cell.SetCellValue(result);
                ICellStyle cellStyle = sheet.Workbook.CreateCellStyle();
                IDataFormat dataFormat = sheet.Workbook.CreateDataFormat();
                cellStyle.DataFormat = dataFormat.GetFormat("#,##0");
                cell.CellStyle = cellStyle;
            }
        }

        /// <summary>
        /// 모든 시트에서 빈 행 제거
        /// </summary>
        /// <param name="sheets">처리할 시트 배열</param>
        private void RemoveEmptyRowsFromAllSheets(ISheet[] sheets)
        {
            foreach (var sheet in sheets)
            {
                RemoveEmptyRows(sheet);
            }
        }

        /// <summary>
        /// 빈 행 제거
        /// </summary>
        /// <param name="sheet">처리할 시트</param>
        private void RemoveEmptyRows(ISheet sheet)
        {
            for (int i = sheet.LastRowNum; i > 0; i--)
            {
                IRow row = sheet.GetRow(i);
                if (row == null || row.Cells.All(d => d.CellType == CellType.Blank))
                {
                    sheet.RemoveRow(row);
                }
            }
        }

        /// <summary>
        /// 워크북 저장
        /// </summary>
        /// <param name="workbook">저장할 워크북</param>
        /// <param name="fileName">파일명</param>
        private void SaveWorkbook(IWorkbook workbook, string fileName)
        {
            string savePath = Path.Combine(outputDirectory, fileName);
            using (FileStream fs = new FileStream(savePath, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fs);
            }
        }

        /// <summary>
        /// 폼 로드 이벤트 핸들러
        /// </summary>
        private void Form1_Load(object sender, EventArgs e)
        {
            // 필요한 초기화 작업 수행
        }
    }

    /// <summary>
    /// 셀 주소 변환 유틸리티 클래스
    /// </summary>
    public static class CellReference
    {
        /// <summary>
        /// 셀 주소를 행과 열 인덱스로 변환
        /// </summary>
        /// <param name="cellReference">셀 주소 (예: "A1")</param>
        /// <returns>행과 열 인덱스 튜플</returns>
        public static (int Row, int Col) ConvertCellReference(string cellReference)
        {
            int row = 0, col = 0;
            foreach (char c in cellReference)
            {
                if (char.IsDigit(c))
                {
                    row = row * 10 + (c - '0');
                }
                else
                {
                    col = col * 26 + (char.ToUpper(c) - 'A' + 1);
                }
            }
            return (row - 1, col - 1); // NPOI는 0부터 인덱스를 사용
        }
    }
}