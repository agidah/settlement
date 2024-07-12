namespace ExcelRowSplitter
{
    partial class Settlement
    {
        private System.ComponentModel.IContainer components = null; // 컴포넌트 컨테이너

        // 리소스를 정리하는 메서드
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose(); // 컴포넌트가 null이 아니면 정리
            }
            base.Dispose(disposing); // 기본 정리 메서드 호출
        }

        // UI 컴포넌트를 초기화하는 메서드
        private void InitializeComponent()
        {
            this.btnAttachFile = new System.Windows.Forms.Button(); // 파일 첨부 버튼
            this.lblFilePath = new System.Windows.Forms.Label(); // 파일 경로 표시 라벨
            this.progressBar = new System.Windows.Forms.ProgressBar(); // 진행바
            this.SuspendLayout();
            // 
            // btnAttachFile
            // 
            this.btnAttachFile.Location = new System.Drawing.Point(12, 12); // 버튼 위치 설정
            this.btnAttachFile.Name = "btnAttachFile"; // 버튼 이름 설정
            this.btnAttachFile.Size = new System.Drawing.Size(125, 23); // 버튼 크기 설정
            this.btnAttachFile.TabIndex = 0; // 탭 인덱스 설정
            this.btnAttachFile.Text = "데이터파일 첨부"; // 버튼 텍스트 설정
            this.btnAttachFile.UseVisualStyleBackColor = true; // 기본 버튼 스타일 사용
            this.btnAttachFile.Click += new System.EventHandler(this.btnAttachFile_Click); // 클릭 이벤트 핸들러 연결
            // 
            // lblFilePath
            // 
            this.lblFilePath.AutoSize = true; // 자동 크기 조정 설정
            this.lblFilePath.Location = new System.Drawing.Point(143, 17); // 라벨 위치 설정
            this.lblFilePath.Name = "lblFilePath"; // 라벨 이름 설정
            this.lblFilePath.Size = new System.Drawing.Size(0, 15); // 라벨 크기 설정
            this.lblFilePath.TabIndex = 1; // 탭 인덱스 설정
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(12, 50); // 진행바 위치 설정
            this.progressBar.Name = "progressBar"; // 진행바 이름 설정
            this.progressBar.Size = new System.Drawing.Size(760, 23); // 진행바 크기 설정
            this.progressBar.TabIndex = 2; // 탭 인덱스 설정
            // 
            // Form1
            // 
            this.ClientSize = new System.Drawing.Size(800, 450); // 폼 크기 설정
            this.Controls.Add(this.progressBar); // 폼에 진행바 추가
            this.Controls.Add(this.lblFilePath); // 폼에 라벨 추가
            this.Controls.Add(this.btnAttachFile); // 폼에 버튼 추가
            this.Name = "Form1"; // 폼 이름 설정
            this.Text = "정산서 발급기"; // 폼 제목 설정
            this.Load += new System.EventHandler(this.Form1_Load); // 폼 로드 이벤트 핸들러 연결
            this.ResumeLayout(false); // 레이아웃 조정
            this.PerformLayout(); // 레이아웃 조정
        }

        private System.Windows.Forms.Button btnAttachFile; // 파일 첨부 버튼
        private System.Windows.Forms.Label lblFilePath; // 파일 경로 표시 라벨
        private System.Windows.Forms.ProgressBar progressBar; // 진행바
    }
}
