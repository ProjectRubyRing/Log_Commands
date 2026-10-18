# ==========================================================================
#  JBoss EAP gc.log 横断検索  (コンテナ x ECSサービス x タスク)
#  検索条件 (FROM～PAUSE_ERROR) を書き換えてから、全体を SSH 端末に貼り付けて実行
#    対象ファイル : gc.log / gc.log.N (ローテーション連番)
#    対応形式     : JDK 9+ (-Xlog:gc*:file=gc.log:time,...) / JDK 8 (-XX:+PrintGCDateStamps)
#    時刻変換     : ログ時刻 [2026-08-20T01:15:30.123+0000] のオフセットを基に JST へ変換
#    色分け       : ERROR (赤) = Full GC / OutOfMemory / 退避失敗 / 停止時間 PAUSE_ERROR 以上
#                   WARN  (黄) = Humongous / Allocation Stall 等 / 停止時間 PAUSE_WARN 以上
# ==========================================================================
FROM='2026-08-20 09:00:00'                      # 検索開始 JST (日付のみ / 秒省略も可)
TO='2026-08-20 18:00:00'                        # 検索終了 JST (空にすると現在まで)
CONTAINER='front|back'                          # 対象コンテナ    (例: back)
SERVICE='intra-web|intra-api|inter-api|sf-api'  # 対象ECSサービス (例: sf-api)
LEVEL=''                                        # レベル絞込      (空=全て / 例: ERROR|WARN)
KEYWORD=''                                      # キーワード正規表現 (空=なし / 例: Pause / 一致箇所を反転表示)
EXCLUDE=''                                      # 除外正規表現    (空=なし / 例: regions:|Metaspace:|User=)
PAUSE_WARN=500                                  # 停止時間 (ms) がこれ以上で WARN  (空=判定しない)
PAUSE_ERROR=1000                                # 停止時間 (ms) がこれ以上で ERROR (空=判定しない)

TZ=JST-9 find /mnt/logs -mindepth 6 -maxdepth 6 -type f \
    -regextype posix-extended \
    -regex ".*/($CONTAINER)/logs/($SERVICE)/mid/[^/]+/gc\.log(\.[0-9]+)?" \
    -newermt "${FROM:-1970-01-01}" -print0 \
| KW="$KEYWORD" EX="$EXCLUDE" xargs -0 -r gawk \
    -v from="$FROM" -v to="$TO" -v lv="$LEVEL" -v pw="$PAUSE_WARN" -v pe="$PAUSE_ERROR" '
    function norm(s, d) { return s substr(d, length(s) + 1) }
    function jst(m,   e) {
        e = mktime(m[1] " " m[2] " " m[3] " " m[4] " " m[5] " 00", 1)
        e -= (m[8] == "-" ? -1 : 1) * (m[9] * 3600 + m[10] * 60)
        return strftime("%Y-%m-%d %H:%M", e + 9 * 3600, 1)
    }
    function pause(s,   p) {
        if (s ~ /Pause/ && match(s, /([0-9]+[.][0-9]+)ms$/, p)) return p[1] + 0
        if (s ~ /\[(Full )?GC[ (]/ && s !~ /concurrent/ && match(s, /, ([0-9]+[.][0-9]+) secs\]/, p)) return p[1] * 1000
        return ""
    }
    function flush(   lvl, ms) {
        if (keep && (kw == "" || buf ~ kw) && (ex == "" || buf !~ ex)) {
            lvl = (ul ~ /^(debug|trace)$/) ? toupper(ul) : "INFO"
            if (ul == "warning" || buf ~ WRN) lvl = "WARN"
            if (ul == "error" || buf ~ ERR) lvl = "ERROR"
            ms = pause(first)
            if (ms != "" && ms >= pe) lvl = "ERROR"
            if (ms != "" && ms >= pw && lvl != "ERROR") lvl = "WARN"
            if (length(lre) == 0 || match(lvl, lre))
                print key "\037" FILENAME "\037" lvl "\037" (ms == "" ? "" : sprintf("%.3f", ms)) "\037" \
                      (ms != "" && first ~ /Pause Full|Full GC/) "\037" buf
        }
        keep = 0
    }
    BEGIN {
        f = norm(from, "0000-01-01 00:00:00.000")
        t = norm(to,   "9999-12-31 23:59:59.999")
        pw = (pw == "") ? 1e18 : pw + 0; pe = (pe == "") ? 1e18 : pe + 0
        ERR = "OutOfMemory|Out [Oo]f [Mm]emory|Pause Full|Full GC|[Tt]o-space [Ee]xhausted|Evacuation Failure|promotion failed|concurrent mode failure"
        WRN = "Humongous Allocation|Allocation Stall|Degenerated GC|[Tt]o-space [Oo]verflow"
        kw = ENVIRON["KW"]; ex = ENVIRON["EX"]
        if (length(lv)) lre = "^(" toupper(lv) ")$"
    }
    BEGINFILE {
        if (length(ERRNO)) { print "[SKIP] " FILENAME " : " ERRNO > "/dev/stderr"; nextfile }
        print "#F\037" FILENAME
        keep = have = bad = 0
    }
    match($0, /^\[?([0-9]{4})-([0-9]{2})-([0-9]{2})T([0-9]{2}):([0-9]{2}):([0-9]{2})[.,]([0-9]{3})([-+])([0-9]{2}):?([0-9]{2})(\]|:)/, m) {
        msg = substr($0, RLENGTH + 1)
        flush(); have = 1
        h = m[1] m[2] m[3] m[4] m[5] m[8] m[9] m[10]
        if ((h in J) == 0) J[h] = jst(m)
        key = J[h] ":" m[6] "." m[7]
        if (key > t) nextfile
        if (key < f) next
        sub(/^(\[[0-9.]+(s|ms|ns)?\])+/, "", msg)      # JDK 9+ : uptime / pid 等の数値デコレータを除去
        sub(/^ *[0-9.]+: /, "", msg)                   # JDK 8  : JVM 起動からの経過秒を除去
        ul = match(msg, /^\[(trace|debug|info|warning|error) *\]/, v) ? v[1] : ""
        if (ul != "") msg = substr(msg, RLENGTH + 1)
        msg = gensub(/^\[([a-z0-9,]+) *\] */, "[\\1] ", 1, msg)
        sub(/^ +/, "", msg); gsub(/\t/, " ", msg)
        first = buf = msg; keep = 1
        next
    }
    keep { buf = buf "\036" $0; next }
    !have && NF { bad++ }
    ENDFILE { flush(); if (bad) print "#U\037" FILENAME "\037" bad }
' \
| LC_ALL=C sort -s -t $'\037' -k1,1 \
| KW="$KEYWORD" EX="$EXCLUDE" gawk \
    -v from="$FROM" -v to="$TO" -v lv="$LEVEL" -v pw="$PAUSE_WARN" -v pe="$PAUSE_ERROR" \
    -v ctr="$CONTAINER" -v svc="$SERVICE" -v cols="${COLUMNS:-$(tput cols)}" '
    function norm(s, d) { return s substr(d, length(s) + 1) }
    function rep(c, n,   s) { s = sprintf("%" n "s", ""); gsub(/ /, c, s); return s }
    function hl(s) { if (length(kw)) gsub(kw, ESC "7m&" ESC "27m", s); return s }
    function num(n, c) { return (n > 0 ? c : "") sprintf("%7d", n) R }
    function opt(v, d) { return v == "" ? d : v }
    function thr(v) { return v == "" ? "(判定なし)" : v " ms 以上" }
    function pc(v) { if (v == "") return ""; v += 0; return v >= pen ? SC["ERROR"] : v >= pwn ? SC["WARN"] : "" }
    function pms(v, n) { return n ? pc(v) sprintf("%11.3f", v) R : sprintf("%11s", "-") }
    function meta(p,   a, n) {
        n = split(p, a, "/")
        C[p] = a[n-5]; S[p] = a[n-3]
        T[p] = length(a[n-1]) == 32 ? substr(a[n-1], 1, 8) : a[n-1]
    }
    BEGIN {
        FS = "\037"; ESC = "\033["; R = ESC "0m"; G = ESC "90m"; sep = G " | " R
        TAG["ERROR"] = ESC "97;41m"; MSG["ERROR"] = ESC "31m"; SC["ERROR"] = ESC "1;31m"
        TAG["WARN"]  = ESC "30;43m"; MSG["WARN"]  = ESC "33m"; SC["WARN"]  = ESC "1;33m"
        TAG["INFO"]  = ESC "32m";    TAG["DEBUG"] = TAG["TRACE"] = G
        CC["front"]  = ESC "36m";    CC["back"]   = ESC "35m"
        kw = ENVIRON["KW"]; ex = ENVIRON["EX"]
        pwn = (pw == "") ? 1e18 : pw + 0; pen = (pe == "") ? 1e18 : pe + 0
        f = norm(from, "0000-01-01 00:00:00.000")
        t = norm(to,   "9999-12-31 23:59:59.999")
        w = (cols + 0 >= 40) ? cols - 1 : 79
        HR = rep("=", w); LR = rep("-", w)
        period = (from == "" ? "(指定なし)" : f) "  ～  " (to == "" ? "(現在まで)" : t)
        print HR
        print "  JBoss EAP gc.log 検索"
        print "    期間 (JST) : " period
        print "    時刻変換   : ログ時刻のオフセット (例: +0000) を基に JST (UTC+9) に変換して表示・ソート"
        print "    対象       : コンテナ=" ctr "   サービス=" svc
        print "    絞り込み   : レベル=" opt(lv, "(全て)") "   キーワード=" opt(kw, "(なし)") "   除外=" opt(ex, "(なし)")
        print "    判定       : " TAG["ERROR"] " ERROR " R " Full GC / OutOfMemory / 退避失敗 / 停止時間 " thr(pe)
        print "                 " TAG["WARN"] " WARN  " R " Humongous / Allocation Stall 等 / 停止時間 " thr(pw)
        re = "^[0-9]{4}-[0-9]{2}-[0-9]{2} [0-9]{2}:[0-9]{2}:[0-9]{2}[.][0-9]{3}$"
        if (match(f, re) == 0 || match(t, re) == 0) {
            print "  " TAG["ERROR"] " FROM / TO の書式が不正です (YYYY-MM-DD HH:MM:SS) " R
            print HR; bad = 1; exit 1
        }
        print LR
        printf "%-23s%s%-5s%s%-9s%s%-8s%s%-5s%s%9s%s%s\n", "TIMESTAMP (JST)", sep, "CTR", sep, "SERVICE", sep, "TASK", sep, "LEVEL", sep, "PAUSE(ms)", sep, "GC LOG"
        print LR
    }
    $1 == "#F" { if (($2 in C) == 0) meta($2); H[$2] += 0; GH[C[$2] " / " S[$2]] += 0; nfile++; next }
    $1 == "#U" { U[$2] = $3; tu += $3; next }
    {
        p = $2; if ((p in C) == 0) meta(p)
        lvl = $3; mc = MSG[lvl]; n = split($6, L, "\036")
        printf "%s%s%s%-5s%s%s%-9s%s%s%-8s%s%s%s%-5s%s%s%s%9s%s%s%s%s%s\n",
            $1, sep, CC[C[p]], C[p], R, sep, S[p], sep, G, T[p], R, sep,
            TAG[lvl], lvl, R, sep, pc($4), $4, R, sep, mc, hl(L[1]), R
        for (i = 2; i <= n; i++) { s = L[i]; gsub(/\t/, "    ", s); printf "%23s%s%s%s%s\n", "", sep, mc, hl(s), R }
        g = C[p] " / " S[p]; H[p]++; GH[g]++; total++
        if (lvl == "ERROR") { HE[p]++; GE[g]++; te++ }
        if (lvl == "WARN")  { HW[p]++; GW[g]++; tw++ }
        if ($4 != "") {
            v = $4 + 0; GN[g]++; HN[p]++; tn++; GT[g] += v; tt += v
            if (v > GM[g]) GM[g] = v
            if (v > HM[p]) HM[p] = v
            if (v > tm) tm = v
            if ($5) { GF[g]++; tf++ }
            PV[NR] = v; PK[NR] = $1; PP[NR] = p; PL[NR] = L[1]
        }
    }
    END {
        if (bad) exit 1
        if (total == 0) print "  " TAG["WARN"] " 該当するログはありません " R
        PROCINFO["sorted_in"] = "@ind_str_asc"
        print HR
        print "  集計   期間 (JST) : " period
        print LR
        printf "  %-24s%7s%7s%7s%8s%7s%11s%12s\n", "CONTAINER / SERVICE", "HITS", "ERROR", "WARN", "PAUSES", "FULL", "MAX(ms)", "TOTAL(ms)"
        for (g in GH) printf "  %-24s%7d%s%s%8d%s%s%12.3f\n", g, GH[g], num(GE[g], MSG["ERROR"]), num(GW[g], MSG["WARN"]),
            GN[g], num(GF[g], MSG["ERROR"]), pms(GM[g], GN[g]), GT[g]
        print LR
        print "  停止時間 TOP 5"
        PROCINFO["sorted_in"] = "@val_num_desc"; i = 0
        for (k in PV) {
            if (++i > 5) break
            p = PP[k]
            printf "   %d. %s%s%s%-5s%s / %-9s / %-8s%s%s%s%s%s\n", i, PK[k], sep, CC[C[p]], C[p], R, S[p], T[p], sep,
                pms(PV[k], 1), " ms", sep, PL[k]
        }
        if (i == 0) print "     (なし)"
        PROCINFO["sorted_in"] = "@ind_str_asc"
        print LR
        printf "  %7s%7s%7s%11s   %s\n", "HITS", "ERROR", "WARN", "MAX(ms)", "FILE (" (nfile + 0) " 件)"
        for (p in H) printf "  %7d%s%s%s   %s%s\n", H[p], num(HE[p], MSG["ERROR"]), num(HW[p], MSG["WARN"]), pms(HM[p], HN[p]), p,
            ((p in U) ? "  " TAG["WARN"] " 時刻解析不可 " U[p] " 行 " R : "")
        print LR
        printf "  %-24s%7d%s%s%8d%s%s%12.3f\n", "TOTAL", total, num(te, MSG["ERROR"]), num(tw, MSG["WARN"]),
            tn, num(tf, MSG["ERROR"]), pms(tm, tn), tt
        if (tu) print "  " TAG["WARN"] " 時刻を解析できない行が " tu " 行あります (-Xlog に time デコレータがあるか確認してください) " R
        print HR
    }
'
