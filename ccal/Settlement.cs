using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Threading.Tasks;
using System.Configuration;

namespace ExcelRowSplitter
{
    public partial class Settlement : Form
    {
        private const int COLUMN_COUNT = 34;
        private const int SHEET_COUNT = 6;
        private string outputDirectory;
        private string[] headerRow;

        public Settlement()
        {
            InitializeComponent();
        }

        private async void btnAttachFile_Click(object sender, EventArgs e)
        {
            string filePath = SelectExcelFile();
            if (string.IsNullOrEmpty(filePath)) return;

            string outputFolder = SelectOutputFolder();
            if (string.IsNullOrEmpty(outputFolder)) return;

            outputDirectory = outputFolder;
            await ProcessExcelFileAsync(filePath);
        }

        private string SelectExcelFile()
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = ConfigurationManager.AppSettings["ExcelFileFilter"],
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

        private async Task ProcessExcelFileAsync(string filePath)
        {
            try
            {
                var progress = new Progress<int>(value => progressBar.Value = value);
                await Task.Run(() => ProcessExcelFile(filePath, progress));
                MessageBox.Show("정산서 데이터추출 완료!");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"오류가 발생했습니다: {ex.Message}");
                // TODO: 로깅 추가
            }
        }

        private void ProcessExcelFile(string filePath, IProgress<int> progress)
        {
            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                IWorkbook workbook = new XSSFWorkbook(fs);
                if (!ValidateWorkbook(workbook)) return;

                ISheet[] sheets = LoadSheets(workbook);
                ProcessSheets(sheets, progress);
            }
        }

        private bool ValidateWorkbook(IWorkbook workbook)
        {
            if (workbook.NumberOfSheets < SHEET_COUNT)
            {
                MessageBox.Show($"엑셀 파일에 시트가 {SHEET_COUNT}개 이상 존재하지 않습니다.");
                return false;
            }
            return true;
        }

        private ISheet[] LoadSheets(IWorkbook workbook)
        {
            return Enumerable.Range(0, SHEET_COUNT)
                             .Select(i => workbook.GetSheetAt(i))
                             .ToArray();
        }

        private void ProcessSheets(ISheet[] sheets, IProgress<int> progress)
        {
            ISheet sheet1 = sheets[0];
            int rowCount = sheet1.PhysicalNumberOfRows;

            headerRow = ExtractHeaderRow(sheet1);

            for (int row = 1; row < rowCount; row++)
            {
                ProcessRow(row, sheets);
                progress.Report((int)((float)row / (rowCount - 1) * 100));
            }
        }

        private string[] ExtractHeaderRow(ISheet sheet)
        {
            IRow header = sheet.GetRow(0);
            return Enumerable.Range(0, COLUMN_COUNT)
                             .Select(col => header.GetCell(col)?.ToString())
                             .ToArray();
        }

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
                // TODO: 로깅 추가
            }
        }

        private string[] ExtractRowData(IRow row)
        {
            return Enumerable.Range(0, COLUMN_COUNT)
                             .Select(col => row.GetCell(col)?.ToString())
                             .ToArray();
        }

        private bool ValidateRowData(string[] rowData)
        {
            return !string.IsNullOrWhiteSpace(rowData[0]) &&
                   !string.IsNullOrWhiteSpace(rowData[1]) &&
                   !string.IsNullOrWhiteSpace(rowData[2]) &&
                   !string.IsNullOrWhiteSpace(rowData[3]) &&
                   !string.IsNullOrWhiteSpace(rowData[5]);
        }

        private string GenerateFileName(string[] rowData)
        {
            string fileName = $"{rowData[0]}~{rowData[1]}_{rowData[5]}_{rowData[3]}_{rowData[2]}.xlsx";
            return CleanFileName(fileName);
        }

        private string CleanFileName(string fileName)
        {
            return Regex.Replace(fileName, @"[\\/:*?""<>|]", string.Empty);
        }

        private void SaveRowToNewExcelFile(string[] rowData, string fileName, ISheet[] sourceSheets)
        {
            using (IWorkbook newWorkbook = new XSSFWorkbook())
            {
                ISheet[] newSheets = CreateNewSheets(newWorkbook);

                string sheet1C2Value = rowData[2];

                AddHeaderAndDataToSheet1(newSheets[0], rowData);
                CopyAndFilterOtherSheets(sourceSheets, newSheets, sheet1C2Value);

                //RemoveSpecificColumns(newSheets);
                ProcessRanges(newSheets);
                RemoveEmptyRowsFromAllSheets(newSheets);

                SaveWorkbook(newWorkbook, fileName);
            }
        }

        private ISheet[] CreateNewSheets(IWorkbook workbook)
        {
            return Enumerable.Range(0, SHEET_COUNT)
                             .Select(i => workbook.CreateSheet($"Sheet{i + 1}"))
                             .ToArray();
        }

        private void AddHeaderAndDataToSheet1(ISheet sheet, string[] rowData)
        {
            IRow headerRowInNewFile = sheet.CreateRow(0);
            IRow newRow = sheet.CreateRow(1);

            for (int col = 0; col < COLUMN_COUNT; col++)
            {
                headerRowInNewFile.CreateCell(col).SetCellValue(headerRow[col]);
                newRow.CreateCell(col).SetCellValue(rowData[col]);
            }
        }

        private void CopyAndFilterOtherSheets(ISheet[] sourceSheets, ISheet[] targetSheets, string filterValue)
        {
            int[] filterColumnIndices = { 2, 3, 3, 3, 0 };
            for (int i = 1; i < SHEET_COUNT; i++)
            {
                CopyAndFilterSheet(sourceSheets[i], targetSheets[i], filterValue, filterColumnIndices[i - 1]);
            }
        }

        private void CopyAndFilterSheet(ISheet sourceSheet, ISheet targetSheet, string compareValue, int compareColumnIndex)
        {
            CopyHeader(sourceSheet, targetSheet);
            CopyFilteredData(sourceSheet, targetSheet, compareValue, compareColumnIndex);
        }

        private void CopyHeader(ISheet sourceSheet, ISheet targetSheet)
        {
            IRow headerRow = sourceSheet.GetRow(0);
            IRow newHeaderRow = targetSheet.CreateRow(0);

            for (int col = 0; col < headerRow.LastCellNum; col++)
            {
                newHeaderRow.CreateCell(col).SetCellValue(headerRow.GetCell(col).ToString());
            }
        }

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

        
       /*
        private void RemoveSpecificColumns(ISheet[] sheets)
        {
            int[] columnsToRemoveSheet1 = {  };
            int[] columnsToRemoveSheet2 = {  };

            RemoveColumns(sheets[0], columnsToRemoveSheet1);
            RemoveColumns(sheets[1], columnsToRemoveSheet2);
        }

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
        */

        private void ProcessRanges(ISheet[] sheets)
        {
            ProcessRange(sheets[0], "G2", "AH");
            ProcessRange(sheets[1], "I2", "AD");
            ProcessRange(sheets[2], "E2", "O");
            ProcessRange(sheets[3], "G2", "G");
            ProcessRange(sheets[4], "G2", "G");
            ProcessRange(sheets[5], "G2", "O");
            ProcessRange(sheets[5], "T2", "Z");
        }

        private void ProcessRange(ISheet sheet, string startCellAddress, string endColumnLetter)
        {
            var (startRow, startColumn) = CellReference.ConvertCellReference(startCellAddress);
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

        private void RemoveEmptyRowsFromAllSheets(ISheet[] sheets)
        {
            foreach (var sheet in sheets)
            {
                RemoveEmptyRows(sheet);
            }
        }

        private void RemoveEmptyRows(ISheet sheet)
        {
            for (int i = sheet.LastRowNum; i >= 0; i--)
            {
                var row = sheet.GetRow(i);
                if (row == null || row.Cells.All(d => d.CellType == CellType.Blank))
                {
                    sheet.RemoveRow(row);
                }
            }
        }

        private void SaveWorkbook(IWorkbook workbook, string fileName)
        {
            string savePath = Path.Combine(outputDirectory, fileName);
            using (FileStream fs = new FileStream(savePath, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fs);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // 필요한 초기화 작업 수행
        }
    }

    public static class CellReference
    {
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
            return (row - 1, col - 1);
        }
    }
}