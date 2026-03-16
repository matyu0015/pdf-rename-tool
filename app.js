// PDF.jsの設定
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

// グローバル変数（PDF名前変更機能）
let availableNames = [];
let pdfData = [];

// グローバル変数（日程整理機能）
let currentData = [];
let currentMergeStartTimeWithDate = false;
let currentMergeDeadlineTimeWithDate = false;
let currentDeadlineTime = '09:00';

// ===== メインタブ切り替え =====
function switchMainTab(tabId) {
    // すべてのタブボタンとコンテンツを非アクティブに
    document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));

    // クリックされたタブをアクティブに
    event.target.classList.add('active');
    document.getElementById('tab-' + tabId).classList.add('active');
}

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
        const pdf = {
            file: file,
            originalName: file.name,
            newName: '',
            data: null  // 後で設定
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

    // PDFのプレビューをレンダリング（Fileオブジェクトから直接読み込み）
    await renderPdfPreview(pdf.file, canvas);
}

// PDFプレビューのレンダリング
async function renderPdfPreview(file, canvas) {
    try {
        // Fileオブジェクトから新しいArrayBufferを読み込む
        const arrayBuffer = await file.arrayBuffer();
        const uint8Array = new Uint8Array(arrayBuffer);

        const loadingTask = pdfjsLib.getDocument({
            data: uint8Array,
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
async function downloadSinglePdf(index) {
    const pdf = pdfData[index];
    if (!pdf.newName) return;

    // Fileオブジェクトから新しいArrayBufferを読み込む
    const arrayBuffer = await pdf.file.arrayBuffer();
    const blob = new Blob([arrayBuffer], { type: 'application/pdf' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = pdf.newName.endsWith('.pdf') ? pdf.newName : `${pdf.newName}.pdf`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);

    console.log(`ダウンロード: ${a.download}, サイズ: ${arrayBuffer.byteLength} bytes`);
}

// 全PDFのZIPダウンロード
document.getElementById('downloadAll').addEventListener('click', async () => {
    const zip = new JSZip();
    let addedCount = 0;

    // 各PDFファイルを処理
    for (const pdf of pdfData) {
        if (pdf.newName) {
            const fileName = pdf.newName.endsWith('.pdf') ? pdf.newName : `${pdf.newName}.pdf`;
            // Fileオブジェクトから新しいArrayBufferを読み込む
            const arrayBuffer = await pdf.file.arrayBuffer();
            zip.file(fileName, arrayBuffer);
            addedCount++;
            console.log(`ZIP追加: ${fileName}, サイズ: ${arrayBuffer.byteLength} bytes`);
        }
    }

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

// ===== 日程整理機能 =====

// サブタブ切り替え（日程整理内のタブ）
function switchSubTab(tab) {
    document.querySelectorAll('.tab-btn-sub').forEach((btn, i) => {
        btn.classList.toggle('active',
            (i === 0 && tab === 'text') ||
            (i === 1 && tab === 'excel') ||
            (i === 2 && tab === 'pdf')
        );
    });
    document.getElementById('tab-text').classList.toggle('active', tab === 'text');
    document.getElementById('tab-excel').classList.toggle('active', tab === 'excel');
    const pdfTab = document.getElementById('tab-pdf');
    if (pdfTab) {
        pdfTab.classList.toggle('active', tab === 'pdf');
    }
}

// Excel取り込み
function onDragOver(e) {
    e.preventDefault();
    document.getElementById('uploadArea').classList.add('dragover');
}

function onDragLeave(e) {
    document.getElementById('uploadArea').classList.remove('dragover');
}

function onDrop(e) {
    e.preventDefault();
    document.getElementById('uploadArea').classList.remove('dragover');
    const file = e.dataTransfer.files[0];
    if (file) readScheduleExcelFile(file);
}

function onFileSelect(e) {
    const file = e.target.files[0];
    if (file) readScheduleExcelFile(file);
}

// PDF取り込み用のドラッグ&ドロップ処理
function onPdfDragOver(e) {
    e.preventDefault();
    document.getElementById('pdfUploadArea').classList.add('dragover');
}

function onPdfDragLeave(e) {
    document.getElementById('pdfUploadArea').classList.remove('dragover');
}

function onPdfDrop(e) {
    e.preventDefault();
    document.getElementById('pdfUploadArea').classList.remove('dragover');
    const file = e.dataTransfer.files[0];
    if (file && file.type === 'application/pdf') {
        readSchedulePdfFile(file);
    } else {
        document.getElementById('pdfUploadResult').textContent = 'PDFファイルを選択してください。';
        document.getElementById('pdfUploadResult').style.color = '#b91c1c';
    }
}

function onPdfFileSelect(e) {
    const file = e.target.files[0];
    if (file) readSchedulePdfFile(file);
}

// PDFファイルからテキストを抽出して日程を解析
async function readSchedulePdfFile(file) {
    try {
        document.getElementById('pdfUploadResult').textContent = 'PDFを読み込み中...';
        document.getElementById('pdfUploadResult').style.color = '#667eea';

        // PDFファイルを読み込み
        const arrayBuffer = await file.arrayBuffer();
        const loadingTask = pdfjsLib.getDocument({
            data: arrayBuffer,
            cMapUrl: 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/cmaps/',
            cMapPacked: true
        });
        const pdf = await loadingTask.promise;

        // 全ページからテキストを抽出
        let allText = '';
        for (let i = 1; i <= pdf.numPages; i++) {
            const page = await pdf.getPage(i);
            const textContent = await page.getTextContent();
            const pageText = textContent.items.map(item => item.str).join(' ');
            allText += pageText + '\n';
        }

        // デバッグ用：抽出したテキストをコンソールに出力
        console.log('=== 抽出されたテキスト ===');
        console.log(allText);
        console.log('========================');

        // 抽出したテキストから日付と時間をペアリング
        const schedules = extractSchedulesFromPdfText(allText);

        // デバッグ用：抽出された日程を出力
        console.log('=== 抽出された日程 ===');
        console.log(schedules);
        console.log('=====================');

        if (schedules.length === 0) {
            document.getElementById('pdfUploadResult').textContent = '日付データが見つかりませんでした。';
            document.getElementById('pdfUploadResult').style.color = '#b91c1c';
            return;
        }

        // テキストエリアに反映
        document.getElementById('dateInput').value = schedules.join('\n');
        document.getElementById('pdfUploadResult').textContent =
            `${schedules.length}件の日程を取り込みました → テキスト欄に反映しました`;
        document.getElementById('pdfUploadResult').style.color = '#059669';

        switchSubTab('text');
        updateCalendarFromInput();

    } catch(err) {
        console.error('PDF読み込みエラー:', err);
        document.getElementById('pdfUploadResult').textContent = '読み込みエラー：' + err.message;
        document.getElementById('pdfUploadResult').style.color = '#b91c1c';
    }
}

// PDFから抽出したテキストを解析して日程リストを作成
function extractSchedulesFromPdfText(text) {
    const schedules = [];

    // 全体を単語に分割（スペースと改行で分割）
    const words = text.split(/[\s\n]+/).filter(w => w.trim());

    let currentDate = null;
    const datePattern1 = /(\d{1,2})月(\d{1,2})日/; // 3月12日形式
    const datePattern2 = /(\d{4})\/(\d{1,2})\/(\d{1,2})/; // 2026/03/12形式
    const timePattern = /^(\d{1,2}):(\d{2})$/;

    console.log('=== 分割された単語 ===');
    console.log(words);
    console.log('=====================');

    for (let i = 0; i < words.length; i++) {
        const word = words[i].trim();
        if (!word) continue;

        // YYYY/MM/DD形式の日付を検出
        const dateMatch2 = word.match(datePattern2);
        if (dateMatch2) {
            const month = parseInt(dateMatch2[2]);
            const day = parseInt(dateMatch2[3]);
            currentDate = `${month}月${day}日`;
            console.log(`日付検出 (YYYY/MM/DD): ${word} → ${currentDate}`);
            continue;
        }

        // M月D日形式の日付を検出
        const dateMatch1 = word.match(datePattern1);
        if (dateMatch1) {
            currentDate = `${dateMatch1[1]}月${dateMatch1[2]}日`;
            console.log(`日付検出 (M月D日): ${currentDate}`);
            continue;
        }

        // 時間を検出（現在の日付と紐付け）
        const timeMatch = word.match(timePattern);
        if (timeMatch && currentDate) {
            const hour = timeMatch[1].padStart(2, '0');
            const minute = timeMatch[2].padStart(2, '0');
            const schedule = `${currentDate} ${hour}:${minute}`;
            schedules.push(schedule);
            console.log(`日程追加: ${schedule}`);
        }
    }

    // 重複を除去
    return [...new Set(schedules)];
}

function readScheduleExcelFile(file) {
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const wb = XLSX.read(data, { type: 'array', cellDates: true });
            const ws = wb.Sheets[wb.SheetNames[0]];
            const extracted = extractDatesFromSheet(ws);

            if (extracted.length === 0) {
                document.getElementById('uploadResult').textContent = '日付データが見つかりませんでした。';
                document.getElementById('uploadResult').style.color = '#b91c1c';
                return;
            }

            document.getElementById('dateInput').value = extracted.join('\n');
            document.getElementById('uploadResult').textContent =
                `${extracted.length}件の日付を取り込みました → テキスト欄に反映しました`;
            document.getElementById('uploadResult').style.color = '#059669';

            switchSubTab('text');
            updateCalendarFromInput();
        } catch(err) {
            document.getElementById('uploadResult').textContent = '読み込みエラー：' + err.message;
            document.getElementById('uploadResult').style.color = '#b91c1c';
        }
    };
    reader.readAsArrayBuffer(file);
}

function extractDatesFromSheet(ws) {
    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
    const grid = [];

    for (let R = range.s.r; R <= range.e.r; R++) {
        grid[R] = [];
        for (let C = range.s.c; C <= range.e.c; C++) {
            const addr = XLSX.utils.encode_cell({ r: R, c: C });
            const cell = ws[addr];
            grid[R][C] = cell ? cellToString(cell) : null;
        }
    }

    const merges = ws['!merges'] || [];
    for (const merge of merges) {
        const val = grid[merge.s.r][merge.s.c];
        for (let R = merge.s.r; R <= merge.e.r; R++) {
            for (let C = merge.s.c; C <= merge.e.c; C++) {
                grid[R][C] = val;
            }
        }
    }

    const dateColCandidates = new Set();
    for (let R = range.s.r; R <= range.e.r; R++) {
        for (let C = range.s.c; C <= range.e.c; C++) {
            const v = grid[R][C];
            if (v && looksLikeDate(v)) dateColCandidates.add(C);
        }
    }

    if (dateColCandidates.size === 0) return [];

    const dateCol = Math.min(...dateColCandidates);
    const timeCol = dateCol + 1;
    const results = [];
    const seen = new Set();

    for (let R = range.s.r; R <= range.e.r; R++) {
        const dateVal = grid[R][dateCol];
        if (!dateVal || !looksLikeDate(dateVal)) continue;

        const timeVal = (timeCol <= range.e.c) ? grid[R][timeCol] : null;
        const timeStr = timeVal && looksLikeTime(timeVal) ? timeVal.trim() : '';

        const key = dateVal + '|' + timeStr;
        if (seen.has(key)) continue;
        seen.add(key);

        results.push(timeStr ? `${dateVal} ${timeStr}` : dateVal);
    }

    return results;
}

function cellToString(cell) {
    if (cell.t === 'd' && cell.v instanceof Date) {
        const d = cell.v;
        const y = d.getFullYear();
        const m = String(d.getMonth() + 1).padStart(2, '0');
        const day = String(d.getDate()).padStart(2, '0');
        return `${y}/${m}/${day}`;
    }
    if (cell.t === 'n' && cell.v > 1) {
        const d = XLSX.SSF.parse_date_code(cell.v);
        if (d && d.y > 1900 && d.m >= 1 && d.m <= 12 && d.d >= 1 && d.d <= 31) {
            return `${d.y}/${String(d.m).padStart(2,'0')}/${String(d.d).padStart(2,'0')}`;
        }
    }
    if (cell.t === 's') return cell.v.trim();
    return String(cell.v ?? '');
}

function looksLikeTime(s) {
    return /\d{1,2}:\d{2}/.test(s);
}

function looksLikeDate(s) {
    return /\d{4}[-\/]\d{1,2}[-\/]\d{1,2}/.test(s) ||
           /\d{1,2}月\d{1,2}日/.test(s) ||
           /^\d{1,2}\/\d{1,2}$/.test(s) ||
           /(令和|平成|昭和)\d+年/.test(s);
}

// 祝日データ (2024〜2026年)
const HOLIDAYS = new Set([
    '2024-01-01','2024-01-08','2024-02-11','2024-02-12','2024-02-23',
    '2024-03-20','2024-04-29','2024-05-03','2024-05-04','2024-05-05','2024-05-06',
    '2024-07-15','2024-08-11','2024-08-12','2024-09-16','2024-09-22','2024-09-23',
    '2024-10-14','2024-11-03','2024-11-04','2024-11-23',
    '2025-01-01','2025-01-13','2025-02-11','2025-02-23','2025-02-24',
    '2025-03-20','2025-04-29','2025-05-03','2025-05-04','2025-05-05','2025-05-06',
    '2025-07-21','2025-08-11','2025-09-15','2025-09-21','2025-09-22','2025-09-23',
    '2025-10-13','2025-11-03','2025-11-23','2025-11-24',
    '2026-01-01','2026-01-12','2026-02-11','2026-02-23',
    '2026-03-20','2026-04-29','2026-05-03','2026-05-04','2026-05-05','2026-05-06',
    '2026-07-20','2026-08-11','2026-09-21','2026-09-22','2026-09-23',
    '2026-10-12','2026-11-03','2026-11-23',
]);

function isHoliday(date) {
    const key = formatDateKey(date);
    return HOLIDAYS.has(key);
}

function isBusinessDay(date) {
    const d = date.getDay();
    return d !== 0 && d !== 6 && !isHoliday(date);
}

function addBusinessDays(date, days) {
    const dir = days >= 0 ? 1 : -1;
    const absD = Math.abs(days);
    let result = new Date(date);
    let count = 0;
    while (count < absD) {
        result.setDate(result.getDate() + dir);
        if (isBusinessDay(result)) count++;
    }
    return result;
}

function formatDate(date) {
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, '0');
    const d = String(date.getDate()).padStart(2, '0');
    return `${y}/${m}/${d}`;
}

function formatDateKey(date) {
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, '0');
    const d = String(date.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
}

const WEEKDAYS_JA = ['日','月','火','水','木','金','土'];
function getDayOfWeek(date) {
    return WEEKDAYS_JA[date.getDay()];
}

// 和暦変換
const GENGO = [
    { name: '令和', start: new Date('2019-05-01'), base: 2018 },
    { name: '平成', start: new Date('1989-01-08'), base: 1988 },
    { name: '昭和', start: new Date('1926-12-25'), base: 1925 },
];

function wareki2year(gengo, nen) {
    for (const g of GENGO) {
        if (gengo.includes(g.name) || gengo === g.name) {
            return g.base + nen;
        }
    }
    return null;
}

// 日程テキスト解析
function parseText(text, defaultYear) {
    const results = [];
    const lines = text.split(/\n/);

    for (const line of lines) {
        const trimmed = line.trim();
        if (!trimmed) continue;
        const parsed = parseLine(trimmed, defaultYear);
        results.push(...parsed);
    }

    const seen = new Set();
    return results.filter(r => {
        const k = r.dateKey + '|' + r.timeRange;
        if (seen.has(k)) return false;
        seen.add(k);
        return true;
    });
}

function parseLine(text, defaultYear) {
    const results = [];
    let normalized = convertWareki(text, defaultYear);
    const segments = splitSegments(normalized);

    for (const seg of segments) {
        const parsed = parseSegment(seg.trim(), defaultYear);
        results.push(...parsed);
    }
    return results;
}

function convertWareki(text, defaultYear) {
    return text.replace(/(令和|平成|昭和)(\d+)年/g, (_, gengo, nen) => {
        const y = wareki2year(gengo, parseInt(nen, 10));
        return y ? `${y}年` : _;
    });
}

function splitSegments(text) {
    const parts = text.split(/(?<=[日)）\d])\s*[,、，・]\s*(?=\d|\s*\d)/);
    return parts;
}

function parseSegment(seg, defaultYear) {
    if (!seg) return [];

    const timeMatch = seg.match(/(\d{1,2}):(\d{2})\s*[〜~\-–]\s*(\d{1,2}):(\d{2})/);
    let timeRange = '';
    if (timeMatch) {
        const sh = timeMatch[1].padStart(2, '0');
        const sm = timeMatch[2].padStart(2, '0');
        const eh = timeMatch[3].padStart(2, '0');
        const em = timeMatch[4].padStart(2, '0');
        timeRange = `${sh}:${sm}〜${eh}:${em}`;
    }

    const date = extractDate(seg, defaultYear);
    if (!date) return [];

    return [{ date, dateKey: formatDateKey(date), timeRange }];
}

function extractDate(seg, defaultYear) {
    let y, m, d;
    let match;

    match = seg.match(/(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
    if (match) {
        y = parseInt(match[1]); m = parseInt(match[2]); d = parseInt(match[3]);
        return makeDate(y, m, d);
    }

    match = seg.match(/(\d{4})年\s*(\d{1,2})月\s*(\d{1,2})日/);
    if (match) {
        y = parseInt(match[1]); m = parseInt(match[2]); d = parseInt(match[3]);
        return makeDate(y, m, d);
    }

    match = seg.match(/(\d{1,2})月\s*(\d{1,2})日/);
    if (match) {
        m = parseInt(match[1]); d = parseInt(match[2]);
        y = guessYear(defaultYear, m, d);
        return makeDate(y, m, d);
    }

    match = seg.match(/^(\d{1,2})\/(\d{1,2})$/);
    if (match) {
        m = parseInt(match[1]); d = parseInt(match[2]);
        y = guessYear(defaultYear, m, d);
        return makeDate(y, m, d);
    }

    match = seg.match(/(?<![\/\d])(\d{1,2})\/(\d{1,2})(?![\d\/])/);
    if (match) {
        m = parseInt(match[1]); d = parseInt(match[2]);
        y = guessYear(defaultYear, m, d);
        return makeDate(y, m, d);
    }

    return null;
}

function makeDate(y, m, d) {
    if (!y || !m || !d) return null;
    const dt = new Date(y, m - 1, d);
    if (dt.getFullYear() !== y || dt.getMonth() !== m - 1 || dt.getDate() !== d) return null;
    return dt;
}

function guessYear(defaultYear, m, d) {
    return defaultYear;
}

// 時間スロット生成
function generateSlots(startTime, endTime, duration, interval) {
    const slots = [];
    const [sh, sm] = startTime.split(':').map(Number);
    const [eh, em] = endTime.split(':').map(Number);
    let cur = sh * 60 + sm;
    const endMin = eh * 60 + em;

    while (cur + duration <= endMin) {
        const s = `${String(Math.floor(cur / 60)).padStart(2, '0')}:${String(cur % 60).padStart(2, '0')}`;
        const e = `${String(Math.floor((cur + duration) / 60)).padStart(2, '0')}:${String((cur + duration) % 60).padStart(2, '0')}`;
        slots.push(`${s}〜${e}`);
        cur += duration + interval;
    }
    return slots;
}

// カレンダー描画
function renderCalendar(selectedKeys) {
    const wrap = document.getElementById('calendarWrap');
    const grid = document.getElementById('calGrid');

    if (!selectedKeys || selectedKeys.size === 0) {
        wrap.classList.remove('active');
        return;
    }

    const months = new Set();
    for (const key of selectedKeys) {
        const [y, m] = key.split('-');
        months.add(`${y}-${m}`);
    }

    const sortedMonths = Array.from(months).sort();
    wrap.classList.add('active');
    grid.innerHTML = sortedMonths.map(ym => buildMonthCalendar(ym, selectedKeys)).join('');
}

function buildMonthCalendar(ym, selectedKeys) {
    const [y, m] = ym.split('-').map(Number);
    const firstDay = new Date(y, m - 1, 1);
    const lastDay = new Date(y, m, 0);
    const startDow = firstDay.getDay();
    const daysInMonth = lastDay.getDate();

    const dowHeaders = ['日','月','火','水','木','金','土'].map((d, i) => {
        const cls = i === 0 ? 'sun' : i === 6 ? 'sat' : '';
        return `<th class="${cls}">${d}</th>`;
    }).join('');

    let cells = '';
    for (let i = 0; i < startDow; i++) {
        cells += '<td class="empty"></td>';
    }

    for (let day = 1; day <= daysInMonth; day++) {
        const date = new Date(y, m - 1, day);
        const dow = date.getDay();
        const key = `${y}-${String(m).padStart(2,'0')}-${String(day).padStart(2,'0')}`;
        const isSel = selectedKeys.has(key);
        const isHol = isHoliday(date);
        const isSun = dow === 0;
        const isSat = dow === 6;

        let dayClass = 'cal-day';
        if (isSel) dayClass += ' selected';
        else if (isHol) dayClass += ' holiday-bg';
        else if (isSun) dayClass += ' sun';
        else if (isSat) dayClass += ' sat';

        const dot = isHol && !isSel ? '<span class="holiday-dot"></span>' : '';
        cells += `<td><span class="${dayClass}">${day}${dot}</span></td>`;

        if (dow === 6 && day < daysInMonth) cells += '</tr><tr>';
    }

    return `
    <div class="cal-month">
        <div class="cal-month-title">${y}年${m}月</div>
        <table class="cal-table">
            <thead><tr>${dowHeaders}</tr></thead>
            <tbody><tr>${cells}</tr></tbody>
        </table>
    </div>`;
}

// UI 制御
function toggleSlot(cb) {
    document.getElementById('slotSection').classList.toggle('active', cb.checked);
}

let debounceTimer = null;

function updateCalendarFromInput() {
    const text = document.getElementById('dateInput').value.trim();
    if (!text) {
        document.getElementById('calendarWrap').classList.remove('active');
        return;
    }
    const defaultYear = parseInt(document.getElementById('defaultYear').value) || new Date().getFullYear();
    const parsed = parseText(text, defaultYear);
    const keys = new Set(parsed.map(p => p.dateKey));
    renderCalendar(keys);
}

function processSchedule() {
    const text = document.getElementById('dateInput').value.trim();
    if (!text) { alert('日程テキストを入力してください'); return; }

    const defaultYear = parseInt(document.getElementById('defaultYear').value) || 2025;
    const nBefore = parseInt(document.getElementById('businessDaysBefore').value) || 0;
    const enableSlot = document.getElementById('enableSlot').checked;
    const mergeStartTimeWithDate = document.getElementById('mergeStartTimeWithDate').checked;
    const mergeDeadlineTimeWithDate = document.getElementById('mergeDeadlineTimeWithDate').checked;
    const deadlineTime = document.getElementById('deadlineTime').value;
    const slotStart = document.getElementById('slotStart').value;
    const slotEnd = document.getElementById('slotEnd').value;
    const duration = parseInt(document.getElementById('duration').value) || 60;
    const interval = parseInt(document.getElementById('interval').value) || 0;

    const parsed = parseText(text, defaultYear);

    if (parsed.length === 0) {
        const alertBox = document.getElementById('alertBox');
        alertBox.style.display = 'block';
        alertBox.textContent = '日程を認識できませんでした。入力形式を確認してください。';
        document.getElementById('resultCard').style.display = 'block';
        document.getElementById('resultBody').innerHTML = '';
        return;
    }

    document.getElementById('alertBox').style.display = 'none';

    const rows = [];
    for (const p of parsed) {
        const deadline = nBefore > 0 ? addBusinessDays(p.date, -nBefore) : null;

        if (enableSlot) {
            let slotsSource;
            if (p.timeRange) {
                const [s, e] = p.timeRange.split('〜');
                slotsSource = generateSlots(s, e, duration, interval);
            } else {
                slotsSource = generateSlots(slotStart, slotEnd, duration, interval);
            }

            for (const slot of slotsSource) {
                rows.push({
                    date: p.date,
                    dateStr: formatDate(p.date),
                    dayOfWeek: getDayOfWeek(p.date),
                    isHoliday: isHoliday(p.date),
                    isWeekend: p.date.getDay() === 0 || p.date.getDay() === 6,
                    timeSlot: slot,
                    deadline: deadline ? formatDate(deadline) : '',
                    deadlineDow: deadline ? getDayOfWeek(deadline) : '',
                });
            }
        } else {
            rows.push({
                date: p.date,
                dateStr: formatDate(p.date),
                dayOfWeek: getDayOfWeek(p.date),
                isHoliday: isHoliday(p.date),
                isWeekend: p.date.getDay() === 0 || p.date.getDay() === 6,
                timeSlot: p.timeRange || '',
                deadline: deadline ? formatDate(deadline) : '',
                deadlineDow: deadline ? getDayOfWeek(deadline) : '',
            });
        }
    }

    currentData = rows;
    currentMergeStartTimeWithDate = mergeStartTimeWithDate;
    currentMergeDeadlineTimeWithDate = mergeDeadlineTimeWithDate;
    currentDeadlineTime = deadlineTime;
    renderTable(rows, nBefore, mergeStartTimeWithDate, mergeDeadlineTimeWithDate, deadlineTime);
    document.getElementById('resultCard').style.display = 'block';

    const keys = new Set(parsed.map(p => p.dateKey));
    renderCalendar(keys);
}

function renderTable(rows, nBefore, mergeStartTime = false, mergeDeadlineTime = false, deadlineTime = '09:00') {
    const showDeadline = nBefore > 0;
    const showSlot = rows.some(r => r.timeSlot);

    const head = document.getElementById('resultHead');
    let cols = ['実施日', '曜日', '備考'];
    // 開始時間を日付と結合しない場合のみ、時間列を追加
    if (showSlot && !mergeStartTime) cols.splice(2, 0, '時間');
    if (showDeadline) cols.push(`締め切り日（${nBefore}営業日前）`, '締め切り曜日');
    head.innerHTML = '<tr>' + cols.map(c => `<th>${c}</th>`).join('') + '</tr>';

    const body = document.getElementById('resultBody');
    body.innerHTML = rows.map(r => {
        const warn = r.isHoliday ? '<span class="badge badge-red">祝日</span>' :
                      r.isWeekend ? '<span class="badge">休日</span>' : '';

        // 開始時間のみを抽出（例: "09:00〜10:00" → "09:00"）
        let startTime = '';
        if (r.timeSlot) {
            const match = r.timeSlot.match(/^(\d{2}:\d{2})/);
            startTime = match ? match[1] : '';
        }

        // 実施日の表示（開始時間を結合する場合）
        const dateDisplay = mergeStartTime && startTime
            ? `${r.dateStr} ${startTime}`
            : r.dateStr;

        let tds = `<td>${dateDisplay}</td><td>${r.dayOfWeek}</td>`;
        // 開始時間を日付と結合しない場合のみ、時間列を表示
        if (showSlot && !mergeStartTime) tds += `<td>${r.timeSlot}</td>`;
        tds += `<td>${warn}</td>`;

        // 締切日の表示（締切時間を結合する場合）
        if (showDeadline) {
            const deadlineDisplay = mergeDeadlineTime && r.deadline
                ? `${r.deadline} ${deadlineTime}`
                : r.deadline;
            tds += `<td>${deadlineDisplay}</td><td>${r.deadlineDow}</td>`;
        }

        return `<tr>${tds}</tr>`;
    }).join('');
}

// Excel出力
function downloadExcel() {
    if (!currentData || currentData.length === 0) return;

    const nBefore = parseInt(document.getElementById('businessDaysBefore').value) || 0;
    // 現在のトグル状態を使用（リアルタイムで反映）
    const mergeStartTime = document.getElementById('mergeStartTimeWithDate').checked;
    const mergeDeadlineTime = document.getElementById('mergeDeadlineTimeWithDate').checked;
    const deadlineTime = document.getElementById('deadlineTime').value;
    const showSlot = currentData.some(r => r.timeSlot);
    const showDeadline = nBefore > 0;

    const headers = ['実施日', '曜日', '備考'];
    // 開始時間を日付と結合しない場合のみ、時間列を追加
    if (showSlot && !mergeStartTime) headers.splice(2, 0, '時間');
    if (showDeadline) headers.push(`締め切り日（${nBefore}営業日前）`, '締め切り曜日');

    const dataRows = currentData.map(r => {
        const warn = r.isHoliday ? '祝日' : r.isWeekend ? '休日' : '';

        // 開始時間のみを抽出（例: "09:00〜10:00" → "09:00"）
        let startTime = '';
        if (r.timeSlot) {
            const match = r.timeSlot.match(/^(\d{2}:\d{2})/);
            startTime = match ? match[1] : '';
        }

        // 実施日の表示（開始時間を結合する場合）
        const dateDisplay = mergeStartTime && startTime
            ? `${r.dateStr} ${startTime}`
            : r.dateStr;

        const row = [dateDisplay, r.dayOfWeek, warn];
        // 開始時間を日付と結合しない場合のみ、時間列を追加
        if (showSlot && !mergeStartTime) row.splice(2, 0, r.timeSlot);

        // 締切日の表示（締切時間を結合する場合）
        if (showDeadline) {
            const deadlineDisplay = mergeDeadlineTime && r.deadline
                ? `${r.deadline} ${deadlineTime}`
                : r.deadline;
            row.push(deadlineDisplay);
            row.push(r.deadlineDow);
        }
        return row;
    });

    const wsData = [headers, ...dataRows];
    const ws = XLSX.utils.aoa_to_sheet(wsData);

    ws['!cols'] = headers.map(h => ({ wch: Math.max(h.length * 2, 14) }));

    const range = XLSX.utils.decode_range(ws['!ref']);
    for (let C = range.s.c; C <= range.e.c; C++) {
        const addr = XLSX.utils.encode_cell({ r: 0, c: C });
        if (!ws[addr]) continue;
        ws[addr].s = {
            fill: { fgColor: { rgb: 'DBEAFE' } },
            font: { bold: true },
            alignment: { horizontal: 'center' }
        };
    }

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '面接日程');

    const today = new Date();
    const fname = `面接日程_${today.getFullYear()}${String(today.getMonth()+1).padStart(2,'0')}${String(today.getDate()).padStart(2,'0')}.xlsx`;
    XLSX.writeFile(wb, fname);
}

// 初期化（DOMContentLoaded）
document.addEventListener('DOMContentLoaded', () => {
    // 日程整理機能の初期化
    const defaultYearInput = document.getElementById('defaultYear');
    if (defaultYearInput) {
        defaultYearInput.value = new Date().getFullYear();
    }

    const dateInput = document.getElementById('dateInput');
    if (dateInput) {
        dateInput.addEventListener('input', () => {
            clearTimeout(debounceTimer);
            debounceTimer = setTimeout(updateCalendarFromInput, 400);
        });
    }

    const defaultYearChangeInput = document.getElementById('defaultYear');
    if (defaultYearChangeInput) {
        defaultYearChangeInput.addEventListener('change', () => {
            updateCalendarFromInput();
        });
    }

    // 時間結合トグルの変更を監視
    const mergeStartTimeToggle = document.getElementById('mergeStartTimeWithDate');
    if (mergeStartTimeToggle) {
        mergeStartTimeToggle.addEventListener('change', updateTableDisplay);
    }

    const mergeDeadlineTimeToggle = document.getElementById('mergeDeadlineTimeWithDate');
    if (mergeDeadlineTimeToggle) {
        mergeDeadlineTimeToggle.addEventListener('change', (e) => {
            // 締切時間フィールドの表示/非表示を切り替え
            const deadlineTimeSection = document.getElementById('deadlineTimeSection');
            if (deadlineTimeSection) {
                deadlineTimeSection.classList.toggle('active', e.target.checked);
            }
            updateTableDisplay();
        });
    }

    const deadlineTimeInput = document.getElementById('deadlineTime');
    if (deadlineTimeInput) {
        deadlineTimeInput.addEventListener('change', updateTableDisplay);
    }
});

// テーブル表示を更新する共通関数
function updateTableDisplay() {
    if (currentData && currentData.length > 0) {
        const nBefore = parseInt(document.getElementById('businessDaysBefore').value) || 0;
        const mergeStartTime = document.getElementById('mergeStartTimeWithDate').checked;
        const mergeDeadlineTime = document.getElementById('mergeDeadlineTimeWithDate').checked;
        const deadlineTime = document.getElementById('deadlineTime').value;
        renderTable(currentData, nBefore, mergeStartTime, mergeDeadlineTime, deadlineTime);
    }
}

// ===== 一括テキスト置換機能 =====

// グローバル変数（テキスト置換機能）
let replaceWorkbook = null;
let replaceCells = [];
let currentCellIndex = 0;
let selectedSheetName = '';
let editMode = 'placeholder'; // 'placeholder' または 'full'

// エクセル/CSVファイルのアップロード処理
document.getElementById('replaceExcelFile')?.addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    try {
        const data = await file.arrayBuffer();
        const isCSV = file.name.toLowerCase().endsWith('.csv');

        // CSVの場合は、文字エンコーディングを自動判定して読み込む
        if (isCSV) {
            let text;
            try {
                // まずUTF-8で試す
                text = new TextDecoder('utf-8').decode(data);
                // 文字化けチェック（�が含まれている場合は失敗）
                if (text.includes('�')) {
                    throw new Error('UTF-8 decode failed');
                }
            } catch (e) {
                // UTF-8で失敗した場合はShift_JISで試す
                try {
                    text = new TextDecoder('shift-jis').decode(data);
                } catch (e2) {
                    // Shift_JISもサポートされていない場合はそのままUTF-8で
                    text = new TextDecoder('utf-8').decode(data);
                }
            }
            replaceWorkbook = XLSX.read(text, { type: 'string' });
        } else {
            replaceWorkbook = XLSX.read(data);
        }

        // シート名のリストを作成
        const sheetSelect = document.getElementById('sheetSelect');
        sheetSelect.innerHTML = '';
        replaceWorkbook.SheetNames.forEach(name => {
            const option = document.createElement('option');
            option.value = name;
            option.textContent = name;
            sheetSelect.appendChild(option);
        });

        const fileType = isCSV ? 'CSVファイル' : 'エクセルファイル';
        const sheetInfo = isCSV ? '' : `（${replaceWorkbook.SheetNames.length}シート）`;
        document.getElementById('replaceExcelStatus').textContent = `✓ ${fileType}を読み込みました${sheetInfo}`;
        document.getElementById('replaceExcelStatus').className = 'status success';
        document.getElementById('sheetSelectionSection').style.display = 'block';

        // 対象列が変更されたら終了行を自動検出
        setupAutoDetectEndRow();

    } catch (error) {
        document.getElementById('replaceExcelStatus').textContent = 'エラー: ' + error.message;
        document.getElementById('replaceExcelStatus').className = 'status error';
    }
});

// 終了行の自動検出設定
function setupAutoDetectEndRow() {
    const targetColumnInput = document.getElementById('targetColumn');
    const sheetSelect = document.getElementById('sheetSelect');

    // イベントリスナーを削除してから再設定（重複を防ぐ）
    const newTargetColumnInput = targetColumnInput.cloneNode(true);
    targetColumnInput.parentNode.replaceChild(newTargetColumnInput, targetColumnInput);

    const newSheetSelect = sheetSelect.cloneNode(true);
    sheetSelect.parentNode.replaceChild(newSheetSelect, sheetSelect);

    // 対象列またはシートが変更されたら自動検出
    newTargetColumnInput.addEventListener('input', detectEndRow);
    newTargetColumnInput.addEventListener('change', detectEndRow);
    newSheetSelect.addEventListener('change', detectEndRow);

    // 初回実行
    detectEndRow();
}

// データ範囲を自動検出して表示
function detectEndRow() {
    if (!replaceWorkbook) return;

    const column = document.getElementById('targetColumn')?.value.trim().toUpperCase();
    const sheetName = document.getElementById('sheetSelect')?.value;

    if (!column || !sheetName) {
        const infoDiv = document.getElementById('autoDetectInfo');
        if (infoDiv) infoDiv.textContent = '';
        return;
    }

    try {
        const sheet = replaceWorkbook.Sheets[sheetName];
        const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1');

        // 列名をインデックスに変換（A=0, B=1, ...）
        const colIndex = XLSX.utils.decode_col(column);

        // 指定列の最初と最後のデータ行を検出
        let firstRow = null;
        let lastRow = null;

        for (let R = 0; R <= range.e.r; R++) {
            const cellAddress = XLSX.utils.encode_cell({ r: R, c: colIndex });
            const cell = sheet[cellAddress];
            if (cell && cell.v !== undefined && cell.v !== null && String(cell.v).trim() !== '') {
                if (firstRow === null) firstRow = R + 1; // 1-indexed
                lastRow = R + 1; // 1-indexed
            }
        }

        const infoDiv = document.getElementById('autoDetectInfo');
        if (!infoDiv) return;

        if (firstRow !== null && lastRow !== null) {
            const dataCount = lastRow - firstRow + 1;
            infoDiv.textContent = `✓ ${column}列に${dataCount}件のデータを検出しました（${firstRow}行目〜${lastRow}行目）`;
            infoDiv.style.color = '#059669';
        } else {
            infoDiv.textContent = `⚠ ${column}列にデータが見つかりませんでした`;
            infoDiv.style.color = '#d97706';
        }

    } catch (error) {
        console.error('データ範囲の自動検出エラー:', error);
        const infoDiv = document.getElementById('autoDetectInfo');
        if (infoDiv) infoDiv.textContent = '';
    }
}

// セル範囲を読み込む
function loadCellsForReplacement() {
    console.log('=== loadCellsForReplacement 開始 ===');

    if (!replaceWorkbook) {
        alert('先にエクセル/CSVファイルをアップロードしてください');
        return;
    }

    // 編集モードは常に'placeholder'（ドラッグ選択方式）
    editMode = 'placeholder';
    console.log('編集モード:', editMode);

    selectedSheetName = document.getElementById('sheetSelect').value;
    const column = document.getElementById('targetColumn').value.trim().toUpperCase();

    console.log('シート:', selectedSheetName);
    console.log('列:', column);

    if (!column) {
        alert('対象列を入力してください');
        return;
    }

    try {
        const sheet = replaceWorkbook.Sheets[selectedSheetName];
        const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1');
        const colIndex = XLSX.utils.decode_col(column);

        console.log('シート範囲:', range);
        console.log('列インデックス:', colIndex);

        // 指定列のすべてのデータを検出
        let firstRow = null;
        let lastRow = null;

        for (let R = 0; R <= range.e.r; R++) {
            const cellAddress = XLSX.utils.encode_cell({ r: R, c: colIndex });
            const cell = sheet[cellAddress];
            if (cell && cell.v !== undefined && cell.v !== null && String(cell.v).trim() !== '') {
                if (firstRow === null) firstRow = R;
                lastRow = R;
            }
        }

        if (firstRow === null || lastRow === null) {
            alert(`${column}列にデータが見つかりませんでした`);
            return;
        }

        console.log('データ範囲:', firstRow, 'から', lastRow);

        replaceCells = [];

        // 指定列の全データを走査
        for (let R = firstRow; R <= lastRow; R++) {
            const cellAddress = XLSX.utils.encode_cell({ r: R, c: colIndex });
            const cell = sheet[cellAddress];
            if (cell && cell.v !== undefined && cell.v !== null && String(cell.v).trim() !== '') {
                const content = String(cell.v);
                replaceCells.push({
                    address: cellAddress,
                    originalContent: content,
                    currentContent: content,
                    row: R,
                    col: colIndex
                });
            }
        }

        console.log('抽出されたセル数:', replaceCells.length);

        if (replaceCells.length === 0) {
            alert('指定列にデータが見つかりませんでした');
            return;
        }

        currentCellIndex = 0;
        console.log('replacementSectionを表示します');
        document.getElementById('replacementSection').style.display = 'block';
        console.log('showCurrentCell()を呼び出します');
        showCurrentCell();

    } catch (error) {
        alert('セル範囲の読み込みエラー: ' + error.message);
    }
}

// 現在のセルを表示
function showCurrentCell() {
    console.log('=== showCurrentCell 開始 ===');
    console.log('currentCellIndex:', currentCellIndex);
    console.log('replaceCells.length:', replaceCells.length);

    if (currentCellIndex >= replaceCells.length) {
        // すべて完了
        console.log('すべて完了しました');
        document.getElementById('replacementSection').style.display = 'none';
        document.getElementById('downloadReplacedSection').style.display = 'block';
        return;
    }

    const cell = replaceCells[currentCellIndex];
    console.log('現在のセル:', cell);

    document.getElementById('replacementProgress').textContent =
        `進捗: ${currentCellIndex + 1} / ${replaceCells.length}`;
    document.getElementById('currentCellAddress').textContent = cell.address;
    document.getElementById('currentCellContent').textContent = cell.currentContent;

    const inputsDiv = document.getElementById('placeholderInputs');
    inputsDiv.innerHTML = '';

    // ドラッグ選択による置換モード
    const instructionDiv = document.createElement('div');
    instructionDiv.style.marginBottom = '15px';
    instructionDiv.style.padding = '12px';
    instructionDiv.style.background = '#eff6ff';
    instructionDiv.style.borderRadius = '6px';
    instructionDiv.style.color = '#1e40af';
    instructionDiv.style.fontSize = '0.9em';
    instructionDiv.innerHTML = '💡 下のテキストから置換したい部分をドラッグして選択し、「選択範囲を追加」ボタンをクリックしてください';
    inputsDiv.appendChild(instructionDiv);

    // 選択可能なテキストエリア
    const selectableDiv = document.createElement('div');
    selectableDiv.id = 'selectable-content';
    selectableDiv.style.background = 'white';
    selectableDiv.style.padding = '12px';
    selectableDiv.style.border = '2px solid #667eea';
    selectableDiv.style.borderRadius = '6px';
    selectableDiv.style.marginBottom = '12px';
    selectableDiv.style.whiteSpace = 'pre-wrap';
    selectableDiv.style.fontSize = '0.95em';
    selectableDiv.style.lineHeight = '1.6';
    selectableDiv.style.userSelect = 'text';
    selectableDiv.style.cursor = 'text';
    selectableDiv.textContent = cell.currentContent;
    inputsDiv.appendChild(selectableDiv);

    // 選択範囲を追加するボタン
    const addButton = document.createElement('button');
    addButton.className = 'btn';
    addButton.textContent = '✨ 選択範囲を追加';
    addButton.style.marginBottom = '15px';
    addButton.style.background = '#667eea';
    addButton.style.color = 'white';
    addButton.onclick = () => addSelectedRange();
    inputsDiv.appendChild(addButton);

    // 置換リストを表示するエリア
    const replacementListDiv = document.createElement('div');
    replacementListDiv.id = 'replacement-list';
    replacementListDiv.style.marginTop = '15px';
    inputsDiv.appendChild(replacementListDiv);

    // 既存の置換リストがあれば表示
    if (!cell.replacements) {
        cell.replacements = [];
    }
    updateReplacementList();
}

// ドラッグ選択範囲を追加
function addSelectedRange() {
    const selection = window.getSelection();
    const selectedText = selection.toString().trim();

    if (!selectedText) {
        alert('テキストを選択してから「選択範囲を追加」ボタンをクリックしてください');
        return;
    }

    const cell = replaceCells[currentCellIndex];
    if (!cell.replacements) {
        cell.replacements = [];
    }

    // 既に同じテキストが登録されているか確認
    const existing = cell.replacements.find(r => r.original === selectedText);
    if (existing) {
        alert(`「${selectedText}」は既に追加されています`);
        return;
    }

    // 置換リストに追加
    cell.replacements.push({
        original: selectedText,
        replacement: ''
    });

    // リストを更新
    updateReplacementList();

    // 選択をクリア
    selection.removeAllRanges();
}

// 置換リストを更新
function updateReplacementList() {
    const cell = replaceCells[currentCellIndex];
    const listDiv = document.getElementById('replacement-list');

    if (!listDiv) return;

    listDiv.innerHTML = '';

    if (!cell.replacements || cell.replacements.length === 0) {
        listDiv.innerHTML = '<p style="color: #999; font-size: 0.9em;">まだ置換対象が追加されていません</p>';
        return;
    }

    cell.replacements.forEach((item, index) => {
        const itemDiv = document.createElement('div');
        itemDiv.style.marginBottom = '15px';
        itemDiv.style.padding = '12px';
        itemDiv.style.background = '#f8f9fa';
        itemDiv.style.borderRadius = '6px';
        itemDiv.style.border = '1px solid #ddd';

        const originalLabel = document.createElement('div');
        originalLabel.style.fontWeight = '600';
        originalLabel.style.marginBottom = '8px';
        originalLabel.style.color = '#667eea';
        originalLabel.textContent = `置換前: "${item.original}"`;
        itemDiv.appendChild(originalLabel);

        const inputContainer = document.createElement('div');
        inputContainer.style.display = 'flex';
        inputContainer.style.gap = '8px';
        inputContainer.style.alignItems = 'center';

        const label = document.createElement('label');
        label.textContent = '置換後:';
        label.style.minWidth = '60px';
        label.style.fontSize = '0.9em';
        inputContainer.appendChild(label);

        const input = document.createElement('input');
        input.type = 'text';
        input.value = item.replacement;
        input.placeholder = '置換後のテキストを入力';
        input.style.flex = '1';
        input.style.padding = '8px';
        input.style.border = '2px solid #ddd';
        input.style.borderRadius = '6px';
        input.onchange = (e) => {
            item.replacement = e.target.value;
        };
        inputContainer.appendChild(input);

        const deleteBtn = document.createElement('button');
        deleteBtn.textContent = '🗑️';
        deleteBtn.className = 'btn';
        deleteBtn.style.padding = '6px 12px';
        deleteBtn.style.background = '#ef4444';
        deleteBtn.style.color = 'white';
        deleteBtn.style.minWidth = '40px';
        deleteBtn.onclick = () => {
            cell.replacements.splice(index, 1);
            updateReplacementList();
        };
        inputContainer.appendChild(deleteBtn);

        itemDiv.appendChild(inputContainer);
        listDiv.appendChild(itemDiv);
    });
}

// プレースホルダーを抽出（旧方式・参考用に残す）
function extractPlaceholders(text) {
    const regex = /\{\{([^}]+)\}\}/g;
    const placeholders = [];
    let match;
    while ((match = regex.exec(text)) !== null) {
        const fullMatch = match[0]; // {{...}} 全体
        if (!placeholders.includes(fullMatch)) {
            placeholders.push(fullMatch);
        }
    }
    return placeholders;
}

// 現在のセルをスキップ
function skipCurrentCell() {
    currentCellIndex++;
    showCurrentCell();
}

// 置換して次へ
function replaceAndNext() {
    const cell = replaceCells[currentCellIndex];
    let newContent = cell.currentContent;

    // ドラッグ選択による置換を実行
    if (cell.replacements && cell.replacements.length > 0) {
        // 各置換対象を処理
        cell.replacements.forEach(item => {
            if (item.replacement) {
                // エスケープして正規表現で置換（全て置換）
                const escapedOriginal = item.original.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
                newContent = newContent.replace(new RegExp(escapedOriginal, 'g'), item.replacement);
            }
        });
    }

    // セル内容を更新
    cell.currentContent = newContent;

    currentCellIndex++;
    showCurrentCell();
}

// 置換後のエクセルをダウンロード
function downloadReplacedExcel() {
    if (!replaceWorkbook) return;

    const sheet = replaceWorkbook.Sheets[selectedSheetName];

    // 置換内容を適用
    replaceCells.forEach(cell => {
        const cellRef = XLSX.utils.encode_cell({ r: cell.row, c: cell.col });
        if (sheet[cellRef]) {
            sheet[cellRef].v = cell.currentContent;
            sheet[cellRef].t = 's'; // 文字列型
        }
    });

    // エクセルファイルを生成
    const wbout = XLSX.write(replaceWorkbook, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([wbout], { type: 'application/octet-stream' });

    // ダウンロード
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `置換済み_${new Date().toISOString().slice(0, 10)}.xlsx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}
