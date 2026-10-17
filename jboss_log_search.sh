# ==========================================================================
#  JBoss EAP server.log 横断検索  (コンテナ x ECSサービス x タスク)
#  検索条件 (FROM～KEYWORD) を書き換えてから、全体を SSH 端末に貼り付けて実行
# ==========================================================================
FROM='2026-08-26 09:00:00'                      # 検索開始 JST (日付のみ / 秒省略も可)
TO='2026-08-26 18:00:00'                        # 検索終了 JST (空にすると現在まで)
LOG_UTC_OFFSET=0                                # ログ時刻のUTCオフセット (UTC出力=0 / JST出力=9)
CONTAINER='front|back'                          # 対象コンテナ    (例: back)
SERVICE='intra-web|intra-api|inter-api|sf-api'  # 対象ECSサービス (例: sf-api)
LEVEL=''                                        # レベル絞込      (空=全て / 例: ERROR|WARN)
KEYWORD=''                                      # キーワード正規表現 (空=なし / 一致箇所を反転表示)

TZ=JST-9 find /mnt/logs -mindepth 6 -maxdepth 6 -type f \
    -regextype posix-extended \
    -regex ".*/($CONTAINER)/logs/($SERVICE)/mid/[^/]+/server\.log(\.[0-9]{4}-[0-9]{2}-[0-9]{2})?" \
    -newermt "${FROM:-1970-01-01}" -print0 \
| KW="$KEYWORD" xargs -0 -r gawk \
    -v from="$FROM" -v to="$TO" -v tz="$LOG_UTC_OFFSET" -v lv="$LEVEL" '
    function norm(s, d) { return s substr(d, length(s) + 1) }
    function jst(h,   t) {
        if (diff == 0) return h
        t = mktime(substr(h, 1, 4) " " substr(h, 6, 2) " " substr(h, 9, 2) " " substr(h, 12, 2) " 00 00", 1)
        return strftime("%Y-%m-%d %H", t + diff * 3600, 1)
    }
    function flush() {
        if (keep && (kw == "" || buf ~ kw)) print key "\037" FILENAME "\037" lvl "\037" buf
        keep = 0
    }
    BEGIN {
        f = norm(from, "0000-01-01 00:00:00.000")
        t = norm(to,   "9999-12-31 23:59:59.999")
        diff = 9 - tz
        kw = ENVIRON["KW"]
        if (length(lv)) lre = "^(" toupper(lv) ")$"
    }
    BEGINFILE {
        if (length(ERRNO)) { print "[SKIP] " FILENAME " : " ERRNO > "/dev/stderr"; nextfile }
        print "#F\037" FILENAME
    }
    /^[0-9]{4}-[0-9]{2}-[0-9]{2} [0-9]{2}:[0-9]{2}:[0-9]{2}[,.][0-9]{3} / {
        flush()
        h = substr($0, 1, 13)
        if ((h in J) == 0) J[h] = jst(h)
        key = J[h] substr($0, 14, 6) "." substr($0, 21, 3)
        if (key > t) nextfile
        if (key < f) next
        lvl = $3
        if (length(lre) && match(lvl, lre) == 0) next
        buf = substr($0, 25); sub(/^[^ ]+ +/, "", buf)
        keep = 1
        next
    }
    keep { buf = buf "\036" $0 }
    ENDFILE { flush() }
' \
| LC_ALL=C sort -s -t $'\037' -k1,1 \
| KW="$KEYWORD" gawk \
    -v from="$FROM" -v to="$TO" -v tz="$LOG_UTC_OFFSET" -v lv="$LEVEL" \
    -v ctr="$CONTAINER" -v svc="$SERVICE" -v cols="${COLUMNS:-$(tput cols)}" '
    function norm(s, d) { return s substr(d, length(s) + 1) }
    function rep(c, n,   s) { s = sprintf("%" n "s", ""); gsub(/ /, c, s); return s }
    function hl(s) { if (length(kw)) gsub(kw, ESC "7m&" ESC "27m", s); return s }
    function num(n, c) { return (n > 0 ? c : "") sprintf("%7d", n) R }
    function meta(p,   a, n) {
        n = split(p, a, "/")
        C[p] = a[n-5]; S[p] = a[n-3]
        T[p] = length(a[n-1]) == 32 ? substr(a[n-1], 1, 8) : a[n-1]
    }
    BEGIN {
        FS = "\037"; ESC = "\033["; R = ESC "0m"; G = ESC "90m"; sep = G " | " R
        TAG["ERROR"] = TAG["FATAL"] = ESC "97;41m"; MSG["ERROR"] = MSG["FATAL"] = ESC "31m"
        TAG["WARN"]  = ESC "30;43m";                MSG["WARN"]  = ESC "33m"
        TAG["INFO"]  = ESC "32m";                   TAG["DEBUG"] = TAG["TRACE"] = G
        CC["front"]  = ESC "36m";                   CC["back"]   = ESC "35m"
        kw = ENVIRON["KW"]
        f = norm(from, "0000-01-01 00:00:00.000")
        t = norm(to,   "9999-12-31 23:59:59.999")
        w = (cols + 0 >= 40) ? cols - 1 : 79
        HR = rep("=", w); LR = rep("-", w)
        period = (from == "" ? "(指定なし)" : f) "  ～  " (to == "" ? "(現在まで)" : t)
        print HR
        print "  JBoss EAP server.log 検索"
        print "    期間 (JST) : " period
        print "    時刻変換   : ログ時刻 UTC" (tz < 0 ? "" : "+") tz " -> JST (UTC+9) に変換して表示・ソート"
        print "    対象       : コンテナ=" ctr "   サービス=" svc
        print "    絞り込み   : レベル=" (lv == "" ? "(全て)" : lv) "   キーワード=" (kw == "" ? "(なし)" : kw)
        re = "^[0-9]{4}-[0-9]{2}-[0-9]{2} [0-9]{2}:[0-9]{2}:[0-9]{2}[.][0-9]{3}$"
        if (match(f, re) == 0 || match(t, re) == 0) {
            print "  " TAG["ERROR"] " FROM / TO の書式が不正です (YYYY-MM-DD HH:MM:SS) " R
            print HR; bad = 1; exit 1
        }
        print LR
        printf "%-23s%s%-5s%s%-9s%s%-8s%s%-5s%s%s\n", "TIMESTAMP (JST)", sep, "CTR", sep, "SERVICE", sep, "TASK", sep, "LEVEL", sep, "MESSAGE"
        print LR
    }
    $1 == "#F" { if (($2 in C) == 0) meta($2); H[$2] += 0; GH[C[$2] " / " S[$2]] += 0; nfile++; next }
    {
        p = $2; if ((p in C) == 0) meta(p)
        lvl = $3; mc = MSG[lvl]; n = split($4, L, "\036")
        printf "%s%s%s%-5s%s%s%-9s%s%s%-8s%s%s%s%-5s%s%s%s%s%s\n", $1, sep, CC[C[p]], C[p], R, sep, S[p], sep, G, T[p], R, sep, TAG[lvl], lvl, R, sep, mc, hl(L[1]), R
        for (i = 2; i <= n; i++) { s = L[i]; gsub(/\t/, "    ", s); printf "%23s%s%s%s%s\n", "", sep, mc, hl(s), R }
        g = C[p] " / " S[p]; H[p]++; GH[g]++; total++
        if (lvl == "ERROR" || lvl == "FATAL") { HE[p]++; GE[g]++; te++ }
        if (lvl == "WARN")                    { HW[p]++; GW[g]++; tw++ }
    }
    END {
        if (bad) exit 1
        if (total == 0) print "  " TAG["WARN"] " 該当するログはありません " R
        PROCINFO["sorted_in"] = "@ind_str_asc"
        print HR
        print "  集計   期間 (JST) : " period
        print LR
        printf "  %-24s%7s%7s%7s\n", "CONTAINER / SERVICE", "HITS", "ERROR", "WARN"
        for (g in GH) printf "  %-24s%7d%s%s\n", g, GH[g], num(GE[g], MSG["ERROR"]), num(GW[g], MSG["WARN"])
        print LR
        printf "  %7s%7s%7s   %s\n", "HITS", "ERROR", "WARN", "FILE (" (nfile + 0) " 件)"
        for (p in H) printf "  %7d%s%s   %s\n", H[p], num(HE[p], MSG["ERROR"]), num(HW[p], MSG["WARN"]), p
        print LR
        printf "  %-24s%7d%s%s\n", "TOTAL", total, num(te, MSG["ERROR"]), num(tw, MSG["WARN"])
        print HR
    }
'
