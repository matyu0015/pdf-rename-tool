// PDF.jsの設定
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

// グローバル変数
let availableNames = [];
let pdfData = [];

// エクセルファイルの処理
document.getElementById('excelFile').addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    try {
        const data = await file.arrayBuffer();
        const workbook = XLSX.read(data);
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });

        // A列（インデックス0）からデータを取得
        availableNames = jsonData
            .map(row => row[0])
            .filter(name => name && name.toString().trim() !== '');

        if (availableNames.length === 0) {
            showStatus('excelStatus', 'A列にデータが見つかりませんでした', 'error');
            return;
        }

        showStatus('excelStatus', `${availableNames.length}件のファイル名を読み込みました`, 'success');
        displayNameList();
        updatePdfSelects();
    } catch (error) {
        showStatus('excelStatus', 'エクセルファイルの読み込みに失敗しました: ' + error.message, 'error');
    }
});

// 名前リストの表示
function displayNameList() {
    const nameListDiv = document.getElementById('nameList');
    nameListDiv.innerHTML = availableNames
        .map((name, index) => `<div class="name-item">${index + 1}. ${name}</div>`)
        .join('');
    nameListDiv.classList.add('active');
}

// PDFファイルの処理
document.getElementById('pdfFiles').addEventListener('change', async (e) => {
    const files = Array.from(e.target.files);
    if (files.length === 0) return;

    pdfData = [];
    document.getElementById('pdfList').innerHTML = '';

    showStatus('pdfStatus', `${files.length}件のPDFを処理中...`, 'success');

    for (const file of files) {
        // PDFデータを読み込み、コピーを作成して保持
        const arrayBuffer = await file.arrayBuffer();
        // Uint8Arrayに変換してからコピーを作成（detachedを防ぐ）
        const uint8Array = new Uint8Array(arrayBuffer);
        const dataForDownload = new Uint8Array(uint8Array);  // ダウンロード用のコピー
        const dataForPreview = new Uint8Array(uint8Array);   // プレビュー用のコピー

        const pdf = {
            file: file,
            originalName: file.name,
            newName: '',
            data: dataForDownload,  // ダウンロード用データ
            previewData: dataForPreview  // プレビュー用データ
        };

        pdfData.push(pdf);
        await renderPdfItem(pdf, pdfData.length - 1);
    }

    showStatus('pdfStatus', `${files.length}件のPDFを読み込みました`, 'success');
    document.getElementById('downloadSection').style.display = 'block';
});

// PDF項目のレンダリング
async function renderPdfItem(pdf, index) {
    const pdfListDiv = document.getElementById('pdfList');

    const itemDiv = document.createElement('div');
    itemDiv.className = 'pdf-item';
    itemDiv.id = `pdf-item-${index}`;

    const headerDiv = document.createElement('div');
    headerDiv.className = 'pdf-header';

    const nameDiv = document.createElement('div');
    nameDiv.className = 'pdf-original-name';
    nameDiv.textContent = `元のファイル名: ${pdf.originalName}`;

    const controlsDiv = document.createElement('div');
    controlsDiv.className = 'pdf-controls';

    const label = document.createElement('label');
    label.textContent = '新しい名前: ';

    const select = document.createElement('select');
    select.id = `select-${index}`;
    select.innerHTML = '<option value="">名前を選択してください</option>';

    availableNames.forEach((name, i) => {
        const option = document.createElement('option');
        option.value = name;
        option.textContent = name;
        select.appendChild(option);
    });

    select.addEventListener('change', (e) => {
        pdf.newName = e.target.value;
        updateDownloadButton(index);
    });

    const downloadBtn = document.createElement('button');
    downloadBtn.className = 'btn btn-download';
    downloadBtn.textContent = 'ダウンロード';
    downloadBtn.disabled = true;
    downloadBtn.id = `download-${index}`;
    downloadBtn.addEventListener('click', () => downloadSinglePdf(index));

    controlsDiv.appendChild(label);
    controlsDiv.appendChild(select);
    controlsDiv.appendChild(downloadBtn);

    headerDiv.appendChild(nameDiv);
    headerDiv.appendChild(controlsDiv);

    const previewDiv = document.createElement('div');
    previewDiv.className = 'pdf-preview';

    const canvas = document.createElement('canvas');
    canvas.className = 'pdf-canvas';
    previewDiv.appendChild(canvas);

    itemDiv.appendChild(headerDiv);
    itemDiv.appendChild(previewDiv);

    pdfListDiv.appendChild(itemDiv);

    // PDFのプレビューをレンダリング（プレビュー用データを使用）
    await renderPdfPreview(pdf.previewData, canvas);
}

// PDFプレビューのレンダリング
async function renderPdfPreview(pdfDataBytes, canvas) {
    try {
        const loadingTask = pdfjsLib.getDocument({
            data: pdfDataBytes,
            cMapUrl: 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/cmaps/',
            cMapPacked: true
        });
        const pdfDoc = await loadingTask.promise;
        const page = await pdfDoc.getPage(1);

        const viewport = page.getViewport({ scale: 1.5 });
        const context = canvas.getContext('2d');
        canvas.height = viewport.height;
        canvas.width = viewport.width;

        const renderContext = {
            canvasContext: context,
            viewport: viewport
        };

        await page.render(renderContext).promise;
        console.log('PDFプレビュー表示成功');
    } catch (error) {
        console.error('PDFレンダリングエラー:', error);
        // エラーメッセージをキャンバスに表示
        const context = canvas.getContext('2d');
        canvas.width = 400;
        canvas.height = 100;
        context.fillStyle = '#f8d7da';
        context.fillRect(0, 0, canvas.width, canvas.height);
        context.fillStyle = '#721c24';
        context.font = '14px Arial';
        context.fillText('PDFの読み込みに失敗しました', 10, 50);
    }
}

// セレクトボックスの更新
function updatePdfSelects() {
    pdfData.forEach((pdf, index) => {
        const select = document.getElementById(`select-${index}`);
        if (select) {
            const currentValue = select.value;
            select.innerHTML = '<option value="">名前を選択してください</option>';

            availableNames.forEach((name) => {
                const option = document.createElement('option');
                option.value = name;
                option.textContent = name;
                select.appendChild(option);
            });

            if (currentValue && availableNames.includes(currentValue)) {
                select.value = currentValue;
            }
        }
    });
}

// ダウンロードボタンの状態更新
function updateDownloadButton(index) {
    const downloadBtn = document.getElementById(`download-${index}`);
    const pdf = pdfData[index];
    downloadBtn.disabled = !pdf.newName;
}

// 単一PDFのダウンロード
function downloadSinglePdf(index) {
    const pdf = pdfData[index];
    if (!pdf.newName) return;

    // Uint8Arrayを使用してBlobを作成
    const blob = new Blob([pdf.data], { type: 'application/pdf' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = pdf.newName.endsWith('.pdf') ? pdf.newName : `${pdf.newName}.pdf`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);

    console.log(`ダウンロード: ${a.download}, サイズ: ${pdf.data.byteLength} bytes`);
}

// 全PDFのZIPダウンロード
document.getElementById('downloadAll').addEventListener('click', async () => {
    const zip = new JSZip();
    let addedCount = 0;

    pdfData.forEach((pdf) => {
        if (pdf.newName) {
            const fileName = pdf.newName.endsWith('.pdf') ? pdf.newName : `${pdf.newName}.pdf`;
            // Uint8Arrayを使用してZIPに追加
            zip.file(fileName, pdf.data);
            addedCount++;
            console.log(`ZIP追加: ${fileName}, サイズ: ${pdf.data.byteLength} bytes`);
        }
    });

    if (addedCount === 0) {
        alert('ダウンロードするPDFがありません。各PDFに新しい名前を選択してください。');
        return;
    }

    console.log(`${addedCount}個のPDFをZIPに追加しました。生成中...`);
    const content = await zip.generateAsync({ type: 'blob' });
    console.log(`ZIP生成完了: ${content.size} bytes`);

    const url = URL.createObjectURL(content);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'renamed_pdfs.zip';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
});

// ステータス表示
function showStatus(elementId, message, type) {
    const statusDiv = document.getElementById(elementId);
    statusDiv.textContent = message;
    statusDiv.className = `status ${type}`;
}
