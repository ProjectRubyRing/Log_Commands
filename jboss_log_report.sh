#!/bin/bash
# ==========================================================================
#  JBoss EAP ログ検索結果 Excel レポート作成  (RHEL 9)
#  jboss_log_search.sh / jboss_access_log_search.sh / jboss_gc_log_search.sh の
#  実行結果を記録した TeraTerm ログを読み込み、Excel (.xlsx) レポートを出力する
#    使い方     : bash jboss_log_report.sh [-o 出力.xlsx] [-e 文字コード] TeraTermログ...
#    出力       : サマリー シート (検索一覧 / 検索条件 / 集計) + 検索ごとの明細シート (1 件 1 行)
#    文字コード : 自動判定 (BOM 付き UTF-8 / UTF-16, UTF-8, CP932)。-e で明示指定も可 (iconv の名前)
#    TeraTerm   : エスケープシーケンス (色) / 行頭タイムスタンプ / プロンプト等は自動で除去
#    必要       : gawk, iconv, python3 (標準ライブラリのみ) または zip
# ==========================================================================
set -u

usage() {
    cat <<'EOF'
使い方 : bash jboss_log_report.sh [-o 出力.xlsx] [-e 文字コード] TeraTermログ...
  -o  出力ファイル (省略時: 入力 1 件 = 入力と同じ場所に <入力名>.xlsx
                            入力 複数 = カレントに jboss_log_report_YYYYMMDD_HHMMSS.xlsx)
  -e  入力の文字コード (省略時: 自動判定 / 例: UTF-8, CP932, UTF-16LE)
EOF
}

OUT=''; ENC=auto
while getopts ':o:e:h' opt; do
    case $opt in
        o) OUT=$OPTARG ;;
        e) ENC=$OPTARG ;;
        h) usage; exit 0 ;;
        *) usage >&2; exit 2 ;;
    esac
done
shift $((OPTIND - 1))
(( $# )) || { usage >&2; exit 2; }

for f; do [[ -f $f && -r $f ]] || { echo "[ERROR] ファイルを読み込めません : $f" >&2; exit 1; }; done
for c in gawk iconv; do command -v $c >/dev/null || { echo "[ERROR] $c が見つかりません" >&2; exit 1; }; done
if   command -v python3 >/dev/null; then ZIPPER=python3
elif command -v zip     >/dev/null; then ZIPPER=zip
else echo "[ERROR] python3 または zip が必要です (dnf install python3)" >&2; exit 1
fi

if [[ -z $OUT ]]; then
    if (( $# == 1 )); then
        base=${1##*/}; name=${base%.*}; OUT="${name:-$base}.xlsx"
        [[ $1 == */* ]] && OUT="${1%/*}/$OUT"
    else
        OUT="jboss_log_report_$(date +%Y%m%d_%H%M%S).xlsx"
    fi
fi
[[ $OUT == /* ]] || OUT="$PWD/$OUT"

WORK=$(mktemp -d) || exit 1
trap 'rm -rf "$WORK"' EXIT
XD="$WORK/xlsx"
mkdir -p "$WORK/in" "$XD/_rels" "$XD/docProps" "$XD/xl/_rels" "$XD/xl/worksheets"

# --- 入力を UTF-8 に変換 (BOM で UTF-8 / UTF-16 を判定、BOM なしは UTF-8 として不正なら CP932) ---
INS=(); SRC=''
for f; do
    o="$WORK/in/${#INS[@]}.txt"; e=$ENC
    if [[ $e == auto ]]; then
        case $(od -An -tx1 -N3 -- "$f" | tr -d ' \n') in
            efbbbf) e=BOM ;;
            fffe*|feff*) e=UTF-16 ;;
            *) if iconv -f UTF-8 -t UTF-8 -- "$f" > "$o" 2>/dev/null; then e=; else e=CP932; fi ;;
        esac
    fi
    case $e in
        '')  ;;
        BOM) tail -c +4 -- "$f" > "$o" ;;
        *)   iconv -f "$e" -t UTF-8 -- "$f" > "$o" || {
                 echo "[ERROR] 文字コード変換に失敗しました ($e) : $f" >&2
                 [[ $e == CP932 ]] && echo "        glibc-gconv-extra が必要な場合があります。-e で文字コードを指定してください" >&2
                 exit 1; } ;;
    esac
    INS+=("$o"); SRC+="$f"$'\n'
done

# --- 解析して xlsx の各パーツ (XML) を出力 ---
LC_ALL=C.UTF-8 XD="$XD" SRC="${SRC%$'\n'}" NOW="$(date '+%Y-%m-%d %H:%M:%S')" NOWZ="$(date -u '+%Y-%m-%dT%H:%M:%SZ')" gawk '
    function trim(s) { sub(/^[ \t]+/, "", s); sub(/[ \t]+$/, "", s); return s }
    function splitsep(s, A, n,   i, p) {             # " | " 区切りで最大 n 項目に分割 (最後の項目は残り全部)
        for (i = 1; i < n && (p = index(s, " | ")); i++) { A[i] = substr(s, 1, p - 1); s = substr(s, p + 3) }
        A[i] = s; return i
    }
    function note(x) { BNV[b, ++BN[b]] = x }
    function open_block(t) {
        b = ++NB; TY[b] = t; BSRC[b] = src; N[b] = EC[b] = WC[b] = 0; st = 1; ss = ""
    }
    function add_row(s,   t, F, n, k, a, q, m) {
        t = TY[b]; n = splitsep(s, F, NC[t])
        for (k = 1; k < n; k++) F[k] = trim(F[k])
        if (N[b] >= MAXROW) { if (!CUT[b]++) note("明細が " MAXROW " 行を超えたため以降を省略しました"); return }
        k = ++N[b]
        R[b, k, 2] = F[1]; R[b, k, 3] = F[2]; R[b, k, 4] = F[3]; R[b, k, 5] = F[4]; R[b, k, 6] = F[5]
        if (t == "server.log") {                        # [カテゴリ] (スレッド) メッセージ
            m = F[6]; R[b, k, 7] = R[b, k, 8] = R[b, k, 10] = ""
            if (match(m, /^\[([^]]*)\] \(([^)]*)\) ?(.*)$/, a)) { R[b, k, 7] = a[1]; R[b, k, 8] = a[2]; m = a[3] }
            R[b, k, 9] = m
        } else if (t == "access_log") {                 # クライアント ... "メソッド URL プロトコル" ステータス バイト数
            m = F[7]; R[b, k, 7] = F[6]; R[b, k, 8] = R[b, k, 9] = R[b, k, 10] = R[b, k, 11] = ""
            if (match(m, /^[^ "]+/)) R[b, k, 8] = substr(m, 1, RLENGTH)
            if (match(m, /"([^"]*)"( +([0-9]{3}|-) +([^ ]+))?/, a)) {
                split(a[1], q, " "); R[b, k, 9] = q[1]; R[b, k, 10] = q[2]; R[b, k, 11] = a[4]
            }
            R[b, k, 12] = m
        } else {                                        # [タグ] GC(ID) メッセージ
            m = F[7]; R[b, k, 7] = F[6]; R[b, k, 8] = R[b, k, 9] = R[b, k, 11] = ""
            if (match(m, /^\[([a-z0-9,]+)\] (.*)$/, a)) { R[b, k, 8] = a[1]; m = a[2] }
            if (match(m, /^GC\(([0-9]+)\) (.*)$/, a))  { R[b, k, 9] = a[1]; m = a[2] }
            R[b, k, 10] = m
        }
        if (F[5] ~ /^(ERROR|FATAL)$/) EC[b]++
        else if (F[5] == "WARN") WC[b]++
    }
    function summary(s,   a, n, i, k, m, P) {
        if (s ~ /^-+$/) { ss = ""; return }
        if (index(s, "  集計 ") == 1) return
        if (index(s, "  CONTAINER / SERVICE ") == 1) { ss = "G"; return }
        if (index(s, "  停止時間 TOP") == 1) { ss = "T"; return }
        if (index(s, "  TOTAL ") == 1) { n = split(s, a, " "); for (i = 2; i <= n; i++) GT[b, i] = a[i]; GTN[b] = n; return }
        if (match(s, /^ +HITS +ERROR +WARN .*FILE \(([0-9]+) 件\)/, m)) { ss = "F"; FCNT[b] = m[1]; return }
        if (s ~ /時刻を解析できない行が/) { note(trim(s)); return }
        if (index(s, "  STATUS ") == 1) { ss = "S"; s = substr(s, 9) }
        if (ss == "G" && s ~ /^  [^ ]+ \/ [^ ]+ /) {   # front / intra-web  HITS ERROR WARN [PAUSES FULL MAX TOTAL]
            n = split(s, a, " "); k = ++GN[b]; G[b, k, 1] = a[1] " / " a[3]
            for (i = 4; i <= n; i++) G[b, k, i - 2] = a[i]
            GC_[b] = n - 2
        } else if (ss == "S") {                         # 200=10  404=2 ...
            while (match(s, /([0-9]{3}|-)=([0-9]+)/, m)) { k = ++SN[b]; SCD[b, k] = m[1]; SCV[b, k] = m[2]; s = substr(s, RSTART + RLENGTH) }
        } else if (ss == "T" && s ~ /^ +[0-9]+\. /) {   # 1. TIMESTAMP | CTR / SERVICE / TASK | PAUSE ms | GC LOG
            splitsep(s, P, 4); k = ++TN[b]
            sub(/^ +[0-9]+\. /, "", P[1]); split(P[2], a, "/"); sub(/ms *$/, "", P[3])
            TP[b, k, 1] = trim(P[1]); TP[b, k, 2] = trim(a[1]); TP[b, k, 3] = trim(a[2]); TP[b, k, 4] = trim(a[3])
            TP[b, k, 5] = trim(P[3]); TP[b, k, 6] = P[4]
        } else if (ss == "F" && match(s, /^ +([0-9]+) +([0-9]+) +([0-9]+) +(([0-9]+[.][0-9]+|-) +)?([^ ].*)$/, m)) {
            k = ++FN[b]; FL[b, k, 1] = m[1]; FL[b, k, 2] = m[2]; FL[b, k, 3] = m[3]; FL[b, k, 4] = m[5]; FL[b, k, 5] = ""
            if (match(m[6], / +時刻解析不可 ([0-9]+) 行/, a)) { FL[b, k, 5] = a[1]; m[6] = substr(m[6], 1, RSTART - 1) }
            FL[b, k, 6] = trim(m[6])
        }
    }

    # ---------------- xlsx 出力 ----------------
    function cl(n,   s) { s = ""; while (n > 0) { s = substr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", (n - 1) % 26 + 1, 1) s; n = int((n - 1) / 26) } return s }
    function esc(s) {
        gsub(/[\001-\010\013-\037\177]/, "", s)
        gsub(/&/, "\\&amp;", s); gsub(/</, "\\&lt;", s); gsub(/>/, "\\&gt;", s); gsub(/"/, "\\&quot;", s)
        return s
    }
    function dw(s,   n) {                               # 表示幅の目安 (全角 = 2)
        sub(/\n.*/, "", s); n = length(s)
        return n > 200 ? 200 : n + gsub(/[^ -~]/, "", s)
    }
    function xdate(v,   e) {                            # "YYYY-MM-DD HH:MM:SS[.mmm]" -> Excel シリアル値
        if (v !~ /^[0-9]{4}-[0-9]{2}-[0-9]{2} [0-9]{2}:[0-9]{2}:[0-9]{2}/) return ""
        e = mktime(substr(v, 1, 4) " " substr(v, 6, 2) " " substr(v, 9, 2) " " substr(v, 12, 2) " " substr(v, 15, 2) " " substr(v, 18, 2), 1)
        if (e == -1) return ""
        return sprintf("%.10f", (e + (substr(v, 20, 1) == "." ? substr(v, 21, 3) / 1000 : 0)) / 86400 + 25569)
    }
    function isnum(v) { return v ~ /^-?[0-9]+([.][0-9]+)?$/ }
    function ro(r, ht) { CR = r; printf "<row r=\"%d\" ht=\"%s\" customHeight=\"1\">", r, ht > OUT }
    function rc() { print "</row>" > OUT }
    function cs(c, v, st) {                             # 文字列セル
        if (v == "") { printf "<c r=\"%s%d\" s=\"%d\"/>", cl(c), CR, st > OUT; return }
        if (length(v) > 32000) v = substr(v, 1, 32000) "\n...(32,000 文字を超えるため省略)"
        printf "<c r=\"%s%d\" s=\"%d\" t=\"inlineStr\"><is><t xml:space=\"preserve\">%s</t></is></c>", cl(c), CR, st, esc(v) > OUT
    }
    function cn(c, v, st) {                             # 数値セル (数値でなければ文字列)
        if (!isnum(v)) { cs(c, v, st); return }
        printf "<c r=\"%s%d\" s=\"%d\"><v>%s</v></c>", cl(c), CR, st, v > OUT
    }
    function cd(c, v, st, alt,   x) { x = xdate(v); if (x == "") cs(c, v, alt); else cn(c, x, st) }
    function cm(c1, c2, v, st,   c) {                   # 結合セル
        cs(c1, v, st); for (c = c1 + 1; c <= c2; c++) cs(c, "", st)
        if (c2 > c1) MG[++NM] = cl(c1) CR ":" cl(c2) CR
    }
    function link(c, txt, st) { cs(c, txt, st); HLR[++NH] = cl(c) CR; HLL[NH] = "\047" txt "\047!A1"; HLT[NH] = txt }
    function sheet_open(idx, frz, nc, W,   i) {
        OUT = XD "/xl/worksheets/sheet" idx ".xml"; NM = NH = 0
        printf "%s<worksheet xmlns=\"%s\" xmlns:r=\"%s\"><sheetPr><pageSetUpPr fitToPage=\"1\"/></sheetPr>", XMLH, NS, NSR > OUT
        printf "<sheetViews><sheetView workbookViewId=\"0\" showGridLines=\"0\"%s>", (idx == 1 ? " tabSelected=\"1\"" : "") > OUT
        if (frz) printf "<pane ySplit=\"%d\" topLeftCell=\"A%d\" activePane=\"bottomLeft\" state=\"frozen\"/><selection pane=\"bottomLeft\" activeCell=\"A%d\" sqref=\"A%d\"/>", frz, frz + 1, frz + 1, frz + 1 > OUT
        printf "</sheetView></sheetViews><sheetFormatPr defaultRowHeight=\"18\" customHeight=\"1\"/><cols>" > OUT
        for (i = 1; i <= nc; i++) printf "<col min=\"%d\" max=\"%d\" width=\"%s\" customWidth=\"1\"/>", i, i, W[i] > OUT
        print "</cols><sheetData>" > OUT
    }
    function sheet_close(af,   i) {
        printf "</sheetData>" > OUT
        if (af != "") printf "<autoFilter ref=\"%s\"/>", af > OUT
        if (NM) { printf "<mergeCells count=\"%d\">", NM > OUT; for (i = 1; i <= NM; i++) printf "<mergeCell ref=\"%s\"/>", MG[i] > OUT; printf "</mergeCells>" > OUT }
        if (NH) {
            printf "<hyperlinks>" > OUT
            for (i = 1; i <= NH; i++) printf "<hyperlink ref=\"%s\" location=\"%s\" display=\"%s\"/>", HLR[i], esc(HLL[i]), esc(HLT[i]) > OUT
            printf "</hyperlinks>" > OUT
        }
        printf "<pageMargins left=\"0.4\" right=\"0.4\" top=\"0.6\" bottom=\"0.6\" header=\"0.3\" footer=\"0.3\"/>" > OUT
        printf "<pageSetup paperSize=\"9\" orientation=\"landscape\" fitToWidth=\"1\" fitToHeight=\"0\"/>" > OUT
        print "<headerFooter><oddFooter>&amp;C&amp;P / &amp;N</oddFooter></headerFooter></worksheet>" > OUT
        close(OUT)
    }
    function cond(b, key,   i, s) {
        s = ""; for (i = 1; i <= BCN[b]; i++) if (BCK[b, i] == key) s = BCV[b, i]
        gsub(/\n/, "   ", s); return s
    }
    function lvst(v, band) {
        if (v ~ /^(ERROR|FATAL)$/) return S_ERR
        if (v == "WARN") return S_WRN
        return (v ~ /^(DEBUG|TRACE)$/ ? S_DIM : S_CTR) + band
    }

    # 明細シート : No / 列見出し / 種別 / 最小幅 / 最大幅
    function write_detail(b,   t, nc, c, k, w, x, y, HD, KD, MN, MX, W, band, v, last) {
        t = TY[b]; nc = split(COLS[t], x, "|")
        for (c = 1; c <= nc; c++) {
            split(x[c], y, ":"); HD[c] = y[1]; KD[c] = y[2]
            w = int(dw(y[1]) * 1.2 + 5.5)               # 見出し (太字) + フィルタボタンが収まる幅
            MN[c] = y[3] > w ? y[3] : w; MX[c] = y[4] > MN[c] ? y[4] : MN[c]; W[c] = MN[c]
        }
        for (c = 1; c <= nc; c++) if (MX[c] > MN[c]) {
            for (k = 1; k <= N[b] && W[c] < MX[c]; k++) if (dw(R[b, k, c]) + 2 > W[c]) W[c] = dw(R[b, k, c]) + 2
            if (W[c] > MX[c]) W[c] = MX[c]
        }
        sheet_open(b + 1, 4, nc, W); last = cl(nc)
        ro(1, 30); cs(1, "No." b "  JBoss EAP " t " 検索結果", S_TITLE); for (c = 2; c <= nc; c++) cs(c, "", S_TITLE); rc()
        ro(2, 18); cs(1, "期間 (JST) : " cond(b, "期間 (JST)") "      対象 : " cond(b, "対象"), S_META); rc()
        ro(3, 18); cs(1, "絞り込み : " cond(b, "絞り込み") "      件数 : " N[b] " 件 (ERROR " EC[b] " / WARN " WC[b] ")      入力ログ : " BSRC[b], S_META); rc()
        ro(4, 24); for (c = 1; c <= nc; c++) cs(c, HD[c], S_HEAD); rc()
        for (k = 1; k <= N[b]; k++) {
            band = (k - 1) % 2; ro(k + 4, 18); cn(1, k, S_CODE + band)
            for (c = 2; c <= nc; c++) {
                v = R[b, k, c]
                if      (KD[c] == "dtms") cd(c, v, S_DTMS + band, S_CTR + band)
                else if (KD[c] == "dts")  cd(c, v, S_DTS + band, S_CTR + band)
                else if (KD[c] == "lvl")  cs(c, v, lvst(v, band))
                else if (KD[c] == "code") cn(c, v, S_CODE + band)
                else if (KD[c] == "int")  cn(c, v, S_INT + band)
                else if (KD[c] == "dec")  cn(c, v, S_DEC + band)
                else cs(c, v, (KD[c] == "ctr" ? S_CTR : S_TEXT) + band)
            }
            rc()
        }
        if (N[b] == 0) { ro(5, 18); cs(1, BERR[b] != "" ? BERR[b] : "該当するログはありません", S_META); rc() }
        sheet_close("A4:" last (N[b] + 4))
        DN[b] = "\047" SHEET[b] "\047!$A$4:$" last "$" (N[b] + 4)
    }

    # サマリー シート (A:余白 B:見出し C～J:数値 K:長い文字列)
    function kv(lab, val,   n, t) {
        t = val; n = gsub(/\n/, "", t) + 1
        ro(R_, n > 1 ? 18 * n : 20); cs(2, lab, S_LABEL); cm(3, 11, val, S_VALUE); rc(); R_++
    }
    function sec(txt,   c) { ro(R_, 26); cs(2, txt, S_SEC); for (c = 3; c <= 11; c++) cs(c, "", S_SEC); rc(); R_++ }
    function cap(txt) { R_++; ro(R_, 20); cs(2, txt, S_CAP); rc(); R_++ }
    function hdr(list,   x, n, i, c, a) {               # "見出し:開始列:終了列|..."
        n = split(list, x, "|"); ro(R_, 22)
        for (i = 1; i <= n; i++) { split(x[i], a, ":"); if (a[3] > a[2]) cm(a[2], a[3], a[1], S_HEAD); else cs(a[2], a[1], S_HEAD) }
        rc(); R_++
    }
    function write_summary(   W, c, b, k, i, band, gc, ac, st, f) {
        W[1] = 2; W[2] = 30; for (c = 3; c <= 10; c++) W[c] = 11; W[11] = 60
        sheet_open(1, 0, 11, W)
        ro(2, 34); cs(2, "JBoss EAP ログ検索レポート", S_TITLE); for (c = 3; c <= 11; c++) cs(c, "", S_TITLE); rc()
        R_ = 4; kv("作成日時", NOW)
        for (i = 1; i <= NSRC; i++) kv("入力ログ", SRC[i])
        kv("検索数", NB " 件")
        R_++; sec("検索一覧")
        hdr("明細シート:2:2|HITS:3:3|ERROR:4:4|WARN:5:5|ファイル数:6:6|ログ種別:7:7|期間 (JST):8:11")
        for (b = 1; b <= NB; b++) {
            band = (b - 1) % 2; ro(R_, 20)
            link(2, SHEET[b], S_LINK + band); cn(3, N[b], S_INT + band); cn(4, EC[b], (EC[b] ? S_INTB : S_INT) + band)
            cn(5, WC[b], S_INT + band); cn(6, FCNT[b], S_INT + band); cs(7, TY[b], S_CTR + band); cm(8, 11, cond(b, "期間 (JST)"), S_TEXT + band)
            rc(); R_++
        }
        for (b = 1; b <= NB; b++) {
            gc = TY[b] == "gc.log"; ac = TY[b] == "access_log"
            R_++; sec("No." b "  " TY[b] " 検索結果")
            ro(R_, 20); cs(2, "明細シート", S_LABEL); link(3, SHEET[b], S_LINK); for (c = 4; c <= 11; c++) cs(c, "", S_LINK)
            MG[++NM] = "C" R_ ":K" R_; rc(); R_++
            kv("入力ログ", BSRC[b])
            for (i = 1; i <= BCN[b]; i++) kv(BCK[b, i], BCV[b, i])
            kv("件数", N[b] " 件  (ERROR " EC[b] " / WARN " WC[b] ")")
            if (BERR[b] != "") kv("エラー", BERR[b])
            for (i = 1; i <= BN[b]; i++) kv("注意", BNV[b, i])
            if (GN[b] || GTN[b]) {
                cap("■ CONTAINER / SERVICE 別")
                hdr("CONTAINER / SERVICE:2:2|HITS:3:3|ERROR:4:4|WARN:5:5" (gc ? "|PAUSES:6:6|FULL:7:7|MAX(ms):8:8|TOTAL(ms):9:9" : ""))
                for (k = 1; k <= GN[b] + 1; k++) {
                    if (k > GN[b] && !GTN[b]) break
                    band = (k - 1) % 2; ro(R_, 20)
                    if (k <= GN[b]) {
                        cs(2, G[b, k, 1], S_TEXT + band)
                        for (c = 2; c <= GC_[b]; c++) cn(c + 1, G[b, k, c], (c >= 7 ? S_DEC : (c == 3 || c == 6) && G[b, k, c] > 0 ? S_INTB : S_INT) + band)
                    } else {
                        cs(2, "TOTAL", S_TTEXT)
                        for (c = 2; c <= GTN[b]; c++) cn(c + 1, GT[b, c], c >= 7 ? S_TDEC : S_TINT)
                    }
                    rc(); R_++
                }
            }
            if (ac && SN[b]) {
                cap("■ STATUS 別")
                hdr("STATUS:2:2|件数:3:3")
                for (k = 1; k <= SN[b]; k++) { band = (k - 1) % 2; ro(R_, 20); cn(2, SCD[b, k], S_CODE + band); cn(3, SCV[b, k], S_INT + band); rc(); R_++ }
            }
            if (gc && TN[b]) {
                cap("■ 停止時間 TOP " TN[b])
                hdr("TIMESTAMP (JST):2:2|CTR:3:3|SERVICE:4:4|TASK:5:5|PAUSE(ms):6:6|GC LOG:7:11")
                for (k = 1; k <= TN[b]; k++) {
                    band = (k - 1) % 2; ro(R_, 20)
                    cd(2, TP[b, k, 1], S_DTMS + band, S_CTR + band); cs(3, TP[b, k, 2], S_CTR + band); cs(4, TP[b, k, 3], S_TEXT + band)
                    cs(5, TP[b, k, 4], S_CTR + band); cn(6, TP[b, k, 5], S_DEC + band); cm(7, 11, TP[b, k, 6], S_TEXT + band)
                    rc(); R_++
                }
            }
            if (FN[b]) {
                cap("■ ファイル別 (" FCNT[b] " 件)")
                f = 6 + gc + (ac || gc)                 # FILE 列の開始位置
                hdr("CONTAINER / SERVICE / TASK:2:2|HITS:3:3|ERROR:4:4|WARN:5:5" (gc ? "|MAX(ms):6:6" : "") (ac || gc ? "|解析不可行:" (f - 1) ":" (f - 1) : "") "|FILE:" f ":11")
                for (k = 1; k <= FN[b]; k++) {
                    band = (k - 1) % 2; ro(R_, 20)
                    i = split(FL[b, k, 6], st, "/")
                    cs(2, i >= 6 ? st[i - 5] " / " st[i - 3] " / " (length(st[i - 1]) == 32 ? substr(st[i - 1], 1, 8) : st[i - 1]) : "", S_TEXT + band)
                    cn(3, FL[b, k, 1], S_INT + band); cn(4, FL[b, k, 2], (FL[b, k, 2] > 0 ? S_INTB : S_INT) + band); cn(5, FL[b, k, 3], S_INT + band)
                    if (gc) cn(6, FL[b, k, 4], S_DEC + band)
                    if (ac || gc) cn(f - 1, FL[b, k, 5], S_INT + band)
                    cm(f, 11, FL[b, k, 6], S_TEXT + band)
                    rc(); R_++
                }
            }
        }
        sheet_close("")
    }

    BEGIN {
        XD = ENVIRON["XD"]; NOW = ENVIRON["NOW"]; NOWZ = ENVIRON["NOWZ"]
        NSRC = split(ENVIRON["SRC"], SRC, "\n")
        NS  = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        NSR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        XMLH = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
        MAXROW = 1000000
        DATA_RE = "^[0-9]{4}-[0-9]{2}-[0-9]{2} [0-9]{2}:[0-9]{2}:[0-9]{2}([.][0-9]{3})? [|] "
        # 出力行の項目数 (最後の項目はメッセージ全体)
        NC["server.log"] = 6; NC["access_log"] = 7; NC["gc.log"] = 7
        # 続き行 (スタックトレース等) を格納する列
        DC["server.log"] = 10; DC["gc.log"] = 11
        COLS["server.log"] = "No:no:7:7|TIMESTAMP (JST):dtms:25:25|CTR:ctr:7:10|SERVICE:txt:11:20|TASK:ctr:11:14|LEVEL:lvl:8:8|CATEGORY:txt:14:45|THREAD:txt:14:30|MESSAGE:txt:40:100|DETAIL (続き行):txt:20:80"
        COLS["access_log"] = "No:no:7:7|TIMESTAMP (JST):dts:21:21|CTR:ctr:7:10|SERVICE:txt:11:20|TASK:ctr:11:14|LEVEL:lvl:8:8|STS:code:7:7|CLIENT:txt:12:20|METHOD:ctr:10:10|URL:txt:30:70|BYTES:int:10:10|ACCESS LOG:txt:40:100"
        COLS["gc.log"]     = "No:no:7:7|TIMESTAMP (JST):dtms:25:25|CTR:ctr:7:10|SERVICE:txt:11:20|TASK:ctr:11:14|LEVEL:lvl:8:8|PAUSE(ms):dec:12:12|TAGS:txt:10:22|GC ID:code:8:8|GC LOG:txt:40:100|DETAIL (続き行):txt:20:80"
        # cellXfs の番号 (styles.xml と対応 / データ用は +1 で縞模様)
        S_TITLE = 1; S_META = 2; S_SEC = 3; S_CAP = 4; S_HEAD = 5; S_LABEL = 6; S_VALUE = 7
        S_TEXT = 8; S_CTR = 10; S_DTMS = 12; S_DTS = 14; S_INT = 16; S_DEC = 18; S_CODE = 20; S_LINK = 22; S_DIM = 24
        S_ERR = 26; S_WRN = 27; S_TTEXT = 28; S_TINT = 29; S_TDEC = 30; S_INTB = 31
    }
    BEGINFILE { src = SRC[++fi]; st = 0 }
    {
        s = $0                                          # TeraTerm ログの制御文字・タイムスタンプを除去
        gsub(/\033\][^\007\033]*(\007|\033\\)/, "", s)
        gsub(/\033\[[0-9;?<=>]*[ -\/]*[@-~]/, "", s)
        gsub(/\033[()*+][0-9A-Za-z]/, "", s)
        gsub(/\033[ -\/]*[0-~]/, "", s)
        gsub(/[\001-\010\013-\037\177]/, "", s)
        sub(/^\[[^]]*[0-9]:[0-9][0-9]:[0-9][0-9][^]]*\] ?/, "", s)
        sub(/[ \t]+$/, "", s)

        if (match(s, /^  JBoss EAP (server\.log|access_log|gc\.log) 検索$/, m)) { open_block(m[1]); next }
        if (!st) next
        if (index(s, "[SKIP] ") == 1) { note("読込スキップ : " substr(s, 8)); next }
        hr = s ~ /^=+$/ && length(s) >= 20
        lr = s ~ /^-+$/ && length(s) >= 20
        if (st == 1) {                                  # 検索条件
            if (hr) st = 0
            else if (lr) st = 2
            else if (s ~ /FROM \/ TO の書式が不正です/) BERR[b] = trim(s)
            else if (s ~ /^    [^ ]/ && (p = index(s, " : "))) { k = ++BCN[b]; BCK[b, k] = trim(substr(s, 1, p)); BCV[b, k] = trim(substr(s, p + 3)) }
            else if (s ~ /^      +[^ ]/ && BCN[b]) BCV[b, BCN[b]] = BCV[b, BCN[b]] "\n" trim(s)
            next
        }
        if (st == 2 || st == 3) {                       # 列見出し
            if (index(s, "TIMESTAMP") == 1) { st = 3; next }
            if (lr && st == 3) { st = 4; next }
            if (hr) { st = 5; next }
            if (s !~ DATA_RE) next
            st = 4
        }
        if (st == 4) {                                  # 明細
            if (hr) st = 5
            else if (s ~ DATA_RE) add_row(s)
            else if (s ~ /^      +[|]( |$)/ && N[b] && (t = DC[TY[b]]) && !CUT[b]) {
                k = N[b]; R[b, k, t] = (R[b, k, t] == "" ? "" : R[b, k, t] "\n") substr(s, index(s, "|") + 2)
            }
            next
        }
        if (hr) { st = 0; next }                        # 集計
        summary(s)
    }
    END {
        if (!NB) { print "[ERROR] JBoss EAP の検索結果が見つかりません (TeraTerm ログを確認してください)" > "/dev/stderr"; exit 3 }
        for (b = 1; b <= NB; b++) SHEET[b] = sprintf(NB > 99 ? "%03d_%s" : "%02d_%s", b, TY[b])
        for (b = 1; b <= NB; b++) write_detail(b)
        write_summary()

        f = XD "/xl/workbook.xml"
        printf "%s<workbook xmlns=\"%s\" xmlns:r=\"%s\"><bookViews><workbookView activeTab=\"0\"/></bookViews><sheets>", XMLH, NS, NSR > f
        printf "<sheet name=\"サマリー\" sheetId=\"1\" r:id=\"rId1\"/>" > f
        for (b = 1; b <= NB; b++) printf "<sheet name=\"%s\" sheetId=\"%d\" r:id=\"rId%d\"/>", esc(SHEET[b]), b + 1, b + 1 > f
        printf "</sheets><definedNames>" > f
        for (b = 1; b <= NB; b++) printf "<definedName name=\"_xlnm._FilterDatabase\" localSheetId=\"%d\" hidden=\"1\">%s</definedName>", b, esc(DN[b]) > f
        for (b = 1; b <= NB; b++) printf "<definedName name=\"_xlnm.Print_Titles\" localSheetId=\"%d\">%s</definedName>", b, esc("\047" SHEET[b] "\047!$4:$4") > f
        print "</definedNames></workbook>" > f; close(f)

        f = XD "/xl/_rels/workbook.xml.rels"
        printf "%s<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">", XMLH > f
        for (i = 1; i <= NB + 1; i++) printf "<Relationship Id=\"rId%d\" Type=\"%s/worksheet\" Target=\"worksheets/sheet%d.xml\"/>", i, NSR, i > f
        printf "<Relationship Id=\"rId%d\" Type=\"%s/styles\" Target=\"styles.xml\"/>", NB + 2, NSR > f
        print "</Relationships>" > f; close(f)

        f = XD "/[Content_Types].xml"
        printf "%s<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">", XMLH > f
        printf "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/>" > f
        printf "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>" > f
        printf "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>" > f
        for (i = 1; i <= NB + 1; i++) printf "<Override PartName=\"/xl/worksheets/sheet%d.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>", i > f
        printf "<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>" > f
        print "</Types>" > f; close(f)

        f = XD "/_rels/.rels"
        printf "%s<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">", XMLH > f
        printf "<Relationship Id=\"rId1\" Type=\"%s/officeDocument\" Target=\"xl/workbook.xml\"/>", NSR > f
        printf "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/>" > f
        print "</Relationships>" > f; close(f)

        f = XD "/docProps/core.xml"
        printf "%s<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">", XMLH > f
        printf "<dc:title>JBoss EAP ログ検索レポート</dc:title><dc:creator>jboss_log_report</dc:creator>" > f
        printf "<dcterms:created xsi:type=\"dcterms:W3CDTF\">%s</dcterms:created><dcterms:modified xsi:type=\"dcterms:W3CDTF\">%s</dcterms:modified>", NOWZ, NOWZ > f
        print "</cp:coreProperties>" > f; close(f)

        # スタイル : Meiryo UI / モノトーン (見出し=濃灰 / 縞=淡灰 / ERROR=黒地白字 / WARN=灰地太字 / DEBUG,TRACE=灰字)
        f = XD "/xl/styles.xml"
        FONT = "<name val=\"Meiryo UI\"/><family val=\"3\"/><charset val=\"128\"/></font>"
        BD = "<left style=\"thin\"><color rgb=\"%s\"/></left><right style=\"thin\"><color rgb=\"%s\"/></right><top style=\"thin\"><color rgb=\"%s\"/></top><bottom style=\"thin\"><color rgb=\"%s\"/></bottom><diagonal/>"
        printf "%s<styleSheet xmlns=\"%s\">", XMLH, NS > f
        printf "<numFmts count=\"3\"><numFmt numFmtId=\"164\" formatCode=\"yyyy-mm-dd hh:mm:ss.000\"/><numFmt numFmtId=\"165\" formatCode=\"yyyy-mm-dd hh:mm:ss\"/><numFmt numFmtId=\"166\" formatCode=\"#,##0.000\"/></numFmts>" > f
        printf "<fonts count=\"8\">" > f
        printf "<font><sz val=\"10\"/><color rgb=\"FF262626\"/>%s", FONT > f            # 0 標準
        printf "<font><b/><sz val=\"10\"/><color rgb=\"FF262626\"/>%s", FONT > f        # 1 太字
        printf "<font><b/><sz val=\"16\"/><color rgb=\"FF000000\"/>%s", FONT > f        # 2 タイトル
        printf "<font><b/><sz val=\"10\"/><color rgb=\"FFFFFFFF\"/>%s", FONT > f        # 3 白太字
        printf "<font><sz val=\"9\"/><color rgb=\"FF595959\"/>%s", FONT > f             # 4 補足
        printf "<font><sz val=\"10\"/><color rgb=\"FF8C8C8C\"/>%s", FONT > f            # 5 淡色
        printf "<font><u/><sz val=\"10\"/><color rgb=\"FF262626\"/>%s", FONT > f        # 6 リンク
        printf "<font><b/><sz val=\"12\"/><color rgb=\"FF000000\"/>%s", FONT > f        # 7 セクション
        printf "</fonts><fills count=\"7\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill>" > f
        split("FF404040 FFF2F2F2 FF000000 FFBFBFBF FFE4E4E4", a, " ")                   # 2 見出し 3 縞 4 ERROR 5 WARN 6 ラベル
        for (i = 1; i <= 5; i++) printf "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"%s\"/><bgColor indexed=\"64\"/></patternFill></fill>", a[i] > f
        printf "</fills><borders count=\"4\"><border><left/><right/><top/><bottom/><diagonal/></border>" > f
        printf "<border>" BD "</border>", "FFBFBFBF", "FFBFBFBF", "FFBFBFBF", "FFBFBFBF" > f   # 1 罫線 (淡)
        printf "<border>" BD "</border>", "FF808080", "FF808080", "FF808080", "FF808080" > f   # 2 罫線 (見出し)
        printf "<border><left/><right/><top/><bottom style=\"medium\"><color rgb=\"FF404040\"/></bottom><diagonal/></border></borders>" > f
        printf "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>" > f
        # numFmt font fill border 横位置 折返し
        nx = split("0 0 0 0 - 0|0 2 0 3 - 0|0 4 0 0 - 0|0 7 0 3 - 0|0 1 0 0 - 0|0 3 2 2 center 1|0 1 6 1 left 0|0 0 0 1 left 1|" \
                   "0 0 0 1 - 0|0 0 3 1 - 0|0 0 0 1 center 0|0 0 3 1 center 0|164 0 0 1 center 0|164 0 3 1 center 0|165 0 0 1 center 0|165 0 3 1 center 0|" \
                   "3 0 0 1 right 0|3 0 3 1 right 0|166 0 0 1 right 0|166 0 3 1 right 0|1 0 0 1 center 0|1 0 3 1 center 0|0 6 0 1 - 0|0 6 3 1 - 0|" \
                   "0 5 0 1 center 0|0 5 3 1 center 0|0 3 4 1 center 0|0 1 5 1 center 0|0 1 6 1 - 0|3 1 6 1 right 0|166 1 6 1 right 0|3 1 0 1 right 0|3 1 3 1 right 0", X, "|")
        printf "<cellXfs count=\"%d\">", nx > f
        for (i = 1; i <= nx; i++) {
            split(X[i], a, " ")
            printf "<xf numFmtId=\"%s\" fontId=\"%s\" fillId=\"%s\" borderId=\"%s\" xfId=\"0\" applyNumberFormat=\"1\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment%s vertical=\"center\"%s/></xf>",
                a[1], a[2], a[3], a[4], (a[5] == "-" ? "" : " horizontal=\"" a[5] "\""), (a[6] ? " wrapText=\"1\"" : "") > f
        }
        print "</cellXfs><cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles></styleSheet>" > f; close(f)

        for (b = 1; b <= NB; b++) { tr += N[b]; msg = msg sprintf("  %-16s %7d 件  (ERROR %d / WARN %d)\n", SHEET[b], N[b], EC[b], WC[b]) }
        printf "検索結果 %d 件 / 明細 %d 行\n%s", NB, tr, msg
    }
' "${INS[@]}" || exit $?

# --- xlsx (zip) にまとめる ---
rm -f -- "$OUT"
if [[ $ZIPPER == python3 ]]; then
    python3 -I - "$XD" "$OUT" <<'PY' || exit 1
import os, sys, zipfile
src, dst = sys.argv[1], sys.argv[2]
with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
    z.write(os.path.join(src, "[Content_Types].xml"), "[Content_Types].xml")
    for root, dirs, files in os.walk(src):
        dirs.sort()
        for f in sorted(files):
            p = os.path.join(root, f)
            a = os.path.relpath(p, src).replace(os.sep, "/")
            if a != "[Content_Types].xml":
                z.write(p, a)
PY
else
    (cd "$XD" && zip -q -X -r "$OUT" .) || exit 1
fi
echo "[OK] Excel レポートを出力しました : $OUT"
