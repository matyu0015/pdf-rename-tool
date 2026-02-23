// PDF.jsの設定
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

// グローバル変数（PDF名前変更機能）
let availableNames = [];
let pdfData = [];

// グローバル変数（日程整理機能）
let currentData = [];

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
        btn.classList.toggle('active', (i === 0 && tab === 'text') || (i === 1 && tab === 'excel'));
    });
    document.getElementById('tab-text').classList.toggle('active', tab === 'text');
    document.getElementById('tab-excel').classList.toggle('active', tab === 'excel');
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
    renderTable(rows, nBefore);
    document.getElementById('resultCard').style.display = 'block';

    const keys = new Set(parsed.map(p => p.dateKey));
    renderCalendar(keys);
}

function renderTable(rows, nBefore) {
    const showDeadline = nBefore > 0;
    const showSlot = rows.some(r => r.timeSlot);

    const head = document.getElementById('resultHead');
    let cols = ['実施日', '曜日', '備考'];
    if (showSlot) cols.splice(2, 0, '時間');
    if (showDeadline) cols.push(`締め切り日（${nBefore}営業日前）`, '締め切り曜日');
    head.innerHTML = '<tr>' + cols.map(c => `<th>${c}</th>`).join('') + '</tr>';

    const body = document.getElementById('resultBody');
    body.innerHTML = rows.map(r => {
        const warn = r.isHoliday ? '<span class="badge badge-red">祝日</span>' :
                      r.isWeekend ? '<span class="badge">休日</span>' : '';
        let tds = `<td>${r.dateStr}</td><td>${r.dayOfWeek}</td>`;
        if (showSlot) tds += `<td>${r.timeSlot}</td>`;
        tds += `<td>${warn}</td>`;
        if (showDeadline) tds += `<td>${r.deadline}</td><td>${r.deadlineDow}</td>`;
        return `<tr>${tds}</tr>`;
    }).join('');
}

// Excel出力
function downloadExcel() {
    if (!currentData || currentData.length === 0) return;

    const nBefore = parseInt(document.getElementById('businessDaysBefore').value) || 0;
    const showSlot = currentData.some(r => r.timeSlot);
    const showDeadline = nBefore > 0;

    const headers = ['実施日', '曜日', '備考'];
    if (showSlot) headers.splice(2, 0, '時間');
    if (showDeadline) headers.push(`締め切り日（${nBefore}営業日前）`, '締め切り曜日');

    const dataRows = currentData.map(r => {
        const warn = r.isHoliday ? '祝日' : r.isWeekend ? '休日' : '';
        const row = [r.dateStr, r.dayOfWeek, warn];
        if (showSlot) row.splice(2, 0, r.timeSlot);
        if (showDeadline) { row.push(r.deadline); row.push(r.deadlineDow); }
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
});
