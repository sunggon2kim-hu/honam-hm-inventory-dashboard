/**
 * 재고 대시보드 데이터 업데이트 스크립트
 *
 * 사용법:
 *   node scripts/update_dashboard.js [--stock <판매재고현황.xlsx>] [--dispose <소진모델.xlsx>] [--branch 호남지사] [--exclude 지점A,지점B]
 *
 * 기본값은 현재 폴더의 판매재고현황_20260915.xlsx / 소진모델.xlsx / 호남지사 이다.
 * 매달 새 원본 파일을 같은 폴더에 넣고 --stock, --dispose 옵션으로 파일명만 바꿔 실행하면
 * dashboard_data.json이 갱신되고, HTML은 그대로 그 데이터를 fetch해서 보여준다.
 */
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const base = path.join(__dirname, '..');

// 폐점 등의 사유로 대시보드에서 제외할 지점 (인도처명 기준). 필요 시 --exclude로 추가/대체 가능.
const DEFAULT_EXCLUDED_BRANCHES = ['2호광장HM', '남악롯데아울렛HM', '상무롯데마트맥스HM', '중화산HM'];

function parseArgs() {
    const args = process.argv.slice(2);
    const opts = {
        stock: '판매재고현황_20260915.xlsx',
        dispose: '소진모델.xlsx',
        branch: '호남지사',
        out: 'dashboard_data.json',
        exclude: DEFAULT_EXCLUDED_BRANCHES,
    };
    for (let i = 0; i < args.length; i++) {
        if (args[i] === '--stock') opts.stock = args[++i];
        else if (args[i] === '--dispose') opts.dispose = args[++i];
        else if (args[i] === '--branch') opts.branch = args[++i];
        else if (args[i] === '--out') opts.out = args[++i];
        else if (args[i] === '--exclude') opts.exclude = args[++i].split(',').map(s => s.trim()).filter(Boolean);
    }
    return opts;
}

function extractCategory(name) {
    const n = String(name).toUpperCase();
    // 세탁기(트롬/워시타워/미니워시/워시콤보/통돌이 등 브랜드명 포함)를 먼저 확인해서
    // "다이얼+LED터치" 같은 표기 때문에 TV로 잘못 분류되는 것을 방지한다.
    if (n.includes('세탁기') || n.includes('WASHER') || n.includes('트롬') ||
        n.includes('워시타워') || n.includes('워시콤보') || n.includes('미니워시') || n.includes('통돌이')) return '세탁기';
    if (n.includes('건조기') || n.includes('DRYER')) return '건조기';
    if (n.includes('식기세척기')) return '식기세척기';
    if (n.includes('TV') || n.includes('OLED') || n.includes('QNED') || n.includes('MRGB') ||
        n.includes('올레드') || n.includes('스탠바이미')) return 'TV';
    if (n.includes('냉장고')) return '냉장고';
    if (n.includes('에어컨') || n.includes('휘센')) return '에어컨';
    if (n.includes('오븐') || n.includes('전자레인지') || n.includes('인덕션') || n.includes('전기레인지')) return '조리가전';
    if (n.includes('공기청정') || n.includes('공청기') || n.includes('에어로타워') || n.includes('알파업')) return '공기청정기';
    if (n.includes('청소기') || n.includes('VACUUM') || n.includes('로보킹') || n.includes('싸이킹')) return '청소기';
    if (n.includes('제습기')) return '제습기';
    return '기타';
}

/**
 * 소진모델.xlsx는 "* 진열" / "* 소진" / "* 스탠바이미" 같은 섹션 제목 행이 섞인
 * 자유 서식 표라서, 헤더 위치를 찾지 않고 [모델명, 상태] 두 칸이 채워진
 * 데이터 행만 골라 사용한다. (컬럼 A는 순번/구분 등으로 섹션마다 형식이 다름)
 */
function loadDisposeMap(filePath) {
    const wb = XLSX.readFile(filePath);
    const map = {};
    for (const sheetName of wb.SheetNames) {
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: null, header: 1 });
        for (const row of rows) {
            const model = row[1];
            const status = row[2];
            if (typeof model === 'string' && (status === '진열대상' || status === '소진대상')) {
                map[model.trim()] = status;
            }
        }
    }
    return map;
}

function build(opts) {
    const stockPath = path.isAbsolute(opts.stock) ? opts.stock : path.join(base, opts.stock);
    const disposePath = path.isAbsolute(opts.dispose) ? opts.dispose : path.join(base, opts.dispose);

    if (!fs.existsSync(stockPath)) throw new Error('재고 파일을 찾을 수 없습니다: ' + stockPath);
    if (!fs.existsSync(disposePath)) throw new Error('소진모델 파일을 찾을 수 없습니다: ' + disposePath);

    const disposeMap = loadDisposeMap(disposePath);

    const wb = XLSX.readFile(stockPath);
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: null });

    const excludeSet = new Set(opts.exclude);
    const filtered = rows.filter(r =>
        r['지사명'] === opts.branch &&
        Number(r['잔여재고']) > 0 &&
        !excludeSet.has(String(r['인도처명']).trim())
    );

    const 전체데이터 = [];
    const 모델별현황 = {};
    const 제품군분포 = {};
    let 진열대상 = 0, 소진대상 = 0, 총재고량 = 0;
    const 지점집합 = new Set();
    let 회전율합 = 0;

    for (const r of filtered) {
        const 상품코드 = String(r['상품코드']).trim();
        const 상품명 = String(r['상품명']).trim();
        const 지점 = String(r['인도처명']).trim();
        const 제품군 = extractCategory(상품명);
        const 진열상태 = disposeMap[상품코드] || '진열대상';
        const 잔여재고 = Number(r['잔여재고']) || 0;
        const 당월판매 = Number(r['당월판매']) || 0;
        const 회전율 = Number(r['회전율']) || 0;

        전체데이터.push({ 지점, 상품명, 상품코드, 제품군, 진열상태, 잔여재고, 당월판매, 회전율 });

        제품군분포[제품군] = (제품군분포[제품군] || 0) + 1;
        if (진열상태 === '소진대상') 소진대상++; else 진열대상++;
        총재고량 += 잔여재고;
        지점집합.add(지점);
        회전율합 += 회전율;

        if (!모델별현황[상품코드]) {
            모델별현황[상품코드] = { 상품명, 제품군, 진열상태, 지점들: [] };
        }
        모델별현황[상품코드].지점들.push({ 지점, 재고: 잔여재고 });
    }

    for (const code in 모델별현황) {
        모델별현황[code].지점들.sort((a, b) => b.재고 - a.재고);
    }

    const 통계 = {
        총재고건: 전체데이터.length,
        소진대상,
        진열대상,
        제품군분포,
        지점수: 지점집합.size,
        총재고량,
        평균회전율: 전체데이터.length ? Number((회전율합 / 전체데이터.length).toFixed(2)) : 0,
    };

    return { 전체데이터, 모델별현황, 통계 };
}

function main() {
    const opts = parseArgs();
    console.log(`재고 파일: ${opts.stock}`);
    console.log(`소진모델 파일: ${opts.dispose}`);
    console.log(`대상 지사: ${opts.branch}`);
    console.log(`제외 지점(폐점 등): ${opts.exclude.join(', ') || '(없음)'}`);

    const data = build(opts);
    const outPath = path.isAbsolute(opts.out) ? opts.out : path.join(base, opts.out);
    fs.writeFileSync(outPath, JSON.stringify(data, null, 0), 'utf-8');

    console.log('\n--- 결과 ---');
    console.log('총재고건:', data.통계.총재고건.toLocaleString());
    console.log('진열대상:', data.통계.진열대상.toLocaleString());
    console.log('소진대상:', data.통계.소진대상.toLocaleString());
    console.log('지점수:', data.통계.지점수);
    console.log('총재고량:', data.통계.총재고량.toLocaleString());
    console.log('평균회전율:', data.통계.평균회전율);
    console.log('제품군분포:', data.통계.제품군분포);
    console.log('\n생성 완료 ->', outPath);
}

main();
