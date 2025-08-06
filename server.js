<!DOCTYPE html>
<html lang="de">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>DB Produktdaten Vergleichstool</title>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: #f5f5f5;
            min-height: 100vh;
            display: flex;
            align-items: center;
            justify-content: center;
            padding: 20px;
        }

        .container {
            max-width: 600px;
            width: 100%;
            background: white;
            border-radius: 8px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.1);
            padding: 40px;
        }

        .header {
            text-align: center;
            margin-bottom: 40px;
        }

        .header h1 {
            font-size: 2em;
            color: #2c3e50;
            font-weight: 300;
            margin-bottom: 8px;
        }

        .header p {
            color: #7f8c8d;
            font-size: 1em;
        }

        .upload-section {
            margin-bottom: 30px;
        }

        .file-upload {
            border: 2px dashed #e74c3c;
            border-radius: 8px;
            padding: 40px 20px;
            text-align: center;
            cursor: pointer;
            transition: all 0.3s ease;
            background: #fafafa;
            margin-bottom: 20px;
        }

        .file-upload:hover {
            background: #f8f9fa;
            border-color: #c0392b;
        }

        .file-upload.has-file {
            border-color: #27ae60;
            background: #f8f9fa;
        }

        .file-upload input {
            display: none;
        }

        .upload-icon {
            font-size: 2.5em;
            margin-bottom: 15px;
            color: #e74c3c;
        }

        .upload-text {
            font-size: 1.1em;
            font-weight: 500;
            margin-bottom: 8px;
            color: #2c3e50;
        }

        .upload-subtext {
            color: #7f8c8d;
            font-size: 0.9em;
        }

        .button-container {
            display: flex;
            gap: 15px;
            justify-content: center;
            margin-bottom: 20px;
        }

        .btn {
            padding: 12px 24px;
            border: none;
            border-radius: 6px;
            cursor: pointer;
            font-size: 14px;
            font-weight: 500;
            transition: all 0.3s ease;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }

        .btn-primary {
            background: #e74c3c;
            color: white;
        }

        .btn-primary:hover {
            background: #c0392b;
        }

        .btn-success {
            background: #27ae60;
            color: white;
        }

        .btn-success:hover {
            background: #229954;
        }

        .btn:disabled {
            background: #bdc3c7;
            cursor: not-allowed;
        }

        .progress-container {
            margin-top: 20px;
            display: none;
        }

        .progress-bar {
            width: 100%;
            height: 6px;
            background: #ecf0f1;
            border-radius: 3px;
            overflow: hidden;
            margin-bottom: 10px;
        }

        .progress-fill {
            height: 100%;
            background: #e74c3c;
            width: 0%;
            transition: width 0.3s ease;
        }

        .progress-text {
            text-align: center;
            color: #2c3e50;
            font-size: 0.9em;
        }

        .status-container {
            margin-top: 20px;
            padding: 12px;
            border-radius: 6px;
            display: none;
            text-align: center;
        }

        .status-success {
            background: #d5f4e6;
            border: 1px solid #27ae60;
            color: #27ae60;
        }

        .status-error {
            background: #fdeaea;
            border: 1px solid #e74c3c;
            color: #e74c3c;
        }

        .color-legend {
            margin-top: 30px;
            padding: 20px;
            background: #f8f9fa;
            border-radius: 6px;
            border-left: 4px solid #e74c3c;
        }

        .color-legend h3 {
            color: #2c3e50;
            margin-bottom: 15px;
            font-size: 1em;
            font-weight: 500;
        }

        .legend-items {
            display: flex;
            gap: 20px;
            flex-wrap: wrap;
        }

        .legend-item {
            display: flex;
            align-items: center;
            gap: 8px;
            font-size: 0.9em;
        }

        .legend-color {
            width: 16px;
            height: 16px;
            border-radius: 3px;
        }

        .legend-green {
            background: #d5f4e6;
            border: 1px solid #27ae60;
        }

        .legend-orange {
            background: #fff3cd;
            border: 1px solid #f39c12;
        }

        .legend-red {
            background: #fdeaea;
            border: 1px solid #e74c3c;
        }

        @media (max-width: 768px) {
            .container {
                padding: 20px;
            }
            
            .button-container {
                flex-direction: column;
            }
            
            .btn {
                width: 100%;
            }
            
            .legend-items {
                flex-direction: column;
                gap: 10px;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>DB Produktdaten Vergleichstool</h1>
            <p>Automatischer Vergleich von Siemens-Produktdaten zwischen Excel-Tabelle und MyMobase</p>
        </div>

        <div class="upload-section">
            <div class="file-upload" id="fileUpload">
                <input type="file" id="excelFile" accept=".xlsx,.xls">
                <div class="upload-icon">📄</div>
                <div class="upload-text" id="uploadText">
                    Excel-Datei hier ablegen oder klicken zum Auswählen
                </div>
                <div class="upload-subtext">
                    Unterstützte Formate: .xlsx, .xls<br>
                    Header in Zeile 3, Daten ab Zeile 4
                </div>
            </div>

            <div class="button-container">
                <button id="processBtn" class="btn btn-primary" disabled>
                    Verarbeiten
                </button>
                <button id="downloadBtn" class="btn btn-success" disabled>
                    Herunterladen
                </button>
            </div>

            <div class="progress-container" id="progressContainer">
                <div class="progress-bar">
                    <div class="progress-fill" id="progressFill"></div>
                </div>
                <div class="progress-text" id="progressText">
                    Bereit zur Verarbeitung...
                </div>
            </div>

            <div class="status-container" id="statusContainer">
                <div id="statusText"></div>
            </div>
        </div>

        <div class="color-legend">
            <h3>Farbmarkierung</h3>
            <div class="legend-items">
                <div class="legend-item">
                    <div class="legend-color legend-green"></div>
                    <span>Grün = Werte stimmen überein</span>
                </div>
                <div class="legend-item">
                    <div class="legend-color legend-orange"></div>
                    <span>Orange = Web-Wert nicht gefunden</span>
                </div>
                <div class="legend-item">
                    <div class="legend-color legend-red"></div>
                    <span>Rot = Abweichung gefunden</span>
                </div>
            </div>
        </div>
    </div>

    <script>
        class ProductComparisonTool {
            constructor() {
                this.excelData = null;
                this.processedFile = null;
                this.initializeEventListeners();
                this.setupFileUpload();
            }

            initializeEventListeners() {
                const processBtn = document.getElementById('processBtn');
                const downloadBtn = document.getElementById('downloadBtn');
                
                processBtn.addEventListener('click', () => this.processFile());
                downloadBtn.addEventListener('click', () => this.downloadFile());
            }

            setupFileUpload() {
                const fileUpload = document.getElementById('fileUpload');
                const fileInput = document.getElementById('excelFile');
                
                fileUpload.addEventListener('click', () => fileInput.click());
                fileUpload.addEventListener('dragover', (e) => {
                    e.preventDefault();
                    fileUpload.style.background = '#f8f9fa';
                });
                fileUpload.addEventListener('dragleave', () => {
                    fileUpload.style.background = '';
                });
                fileUpload.addEventListener('drop', (e) => {
                    e.preventDefault();
                    fileUpload.style.background = '';
                    if (e.dataTransfer.files.length > 0) {
                        this.handleFileUpload(e.dataTransfer.files[0]);
                    }
                });
                
                fileInput.addEventListener('change', (e) => {
                    if (e.target.files.length > 0) {
                        this.handleFileUpload(e.target.files[0]);
                    }
                });
            }

            handleFileUpload(file) {
                if (!file.name.match(/\.(xlsx|xls)$/)) {
                    this.showStatus('Bitte wählen Sie eine gültige Excel-Datei (.xlsx oder .xls)', 'error');
                    return;
                }

                this.excelData = file;
                this.updateFileUploadUI(file);
                this.showStatus(`Datei "${file.name}" erfolgreich geladen`, 'success');
                
                document.getElementById('processBtn').disabled = false;
                document.getElementById('downloadBtn').disabled = true;
            }

            updateFileUploadUI(file) {
                const fileUpload = document.getElementById('fileUpload');
                const uploadText = document.getElementById('uploadText');
                
                fileUpload.classList.add('has-file');
                uploadText.innerHTML = `${file.name}<br><small>${(file.size / 1024 / 1024).toFixed(2)} MB</small>`;
            }

            async processFile() {
                if (!this.excelData) {
                    this.showStatus('Bitte laden Sie zuerst eine Excel-Datei hoch', 'error');
                    return;
                }

                const processBtn = document.getElementById('processBtn');
                const progressContainer = document.getElementById('progressContainer');
                const progressFill = document.getElementById('progressFill');
                const progressText = document.getElementById('progressText');

                processBtn.disabled = true;
                progressContainer.style.display = 'block';
                this.showStatus('Verarbeitung läuft...', 'success');

                try {
                    progressFill.style.width = '25%';
                    progressText.textContent = 'Datei wird hochgeladen...';

                    const formData = new FormData();
                    formData.append('file', this.excelData);

                    progressFill.style.width = '50%';
                    progressText.textContent = 'Web-Scraping läuft...';

                    const response = await fetch('/api/process-excel', {
                        method: 'POST',
                        body: formData
                    });

                    if (!response.ok) {
                        throw new Error(`HTTP ${response.status}: ${response.statusText}`);
                    }

                    progressFill.style.width = '75%';
                    progressText.textContent = 'Verarbeitung abgeschlossen...';

                    const blob = await response.blob();
                    this.processedFile = blob;

                    progressFill.style.width = '100%';
                    progressText.textContent = 'Verarbeitung erfolgreich abgeschlossen!';

                    this.showStatus('Verarbeitung erfolgreich! Sie können jetzt die Datei herunterladen.', 'success');
                    
                    document.getElementById('downloadBtn').disabled = false;

                } catch (error) {
                    console.error('Verarbeitungsfehler:', error);
                    this.showStatus(`Fehler bei der Verarbeitung: ${error.message}`, 'error');
                    progressText.textContent = 'Fehler aufgetreten';
                } finally {
                    processBtn.disabled = false;
                }
            }

            downloadFile() {
                if (!this.processedFile) {
                    this.showStatus('Keine verarbeitete Datei verfügbar', 'error');
                    return;
                }

                const url = URL.createObjectURL(this.processedFile);
                const a = document.createElement('a');
                a.href = url;
                a.download = 'DB_Produktvergleich_verarbeitet.xlsx';
                document.body.appendChild(a);
                a.click();
                a.remove();
                
                setTimeout(() => URL.revokeObjectURL(url), 2000);
                
                this.showStatus('Download erfolgreich gestartet!', 'success');
            }

            showStatus(message, type) {
                const statusContainer = document.getElementById('statusContainer');
                const statusText = document.getElementById('statusText');
                
                statusContainer.className = `status-container status-${type}`;
                statusText.textContent = message;
                statusContainer.style.display = 'block';
                
                if (type === 'success') {
                    setTimeout(() => {
                        statusContainer.style.display = 'none';
                    }, 5000);
                }
            }
        }

        document.addEventListener('DOMContentLoaded', () => {
            new ProductComparisonTool();
        });
    </script>
</body>
</html> 
