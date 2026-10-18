# ==========================================================================
#  JBoss EAP access_log 横断検索  (コンテナ x ECSサービス x タスク)
#  検索条件 (FROM～EXCLUDE) を書き換えてから、全体を SSH 端末に貼り付けて実行
#    対象ファイル : access_log.log / access_log.YYYY-MM-DD.log
#    時刻変換     : ログ時刻 [20/Aug/2026:01:15:30 +0000] のオフセットを基に JST へ変換
#    色分け       : ERROR = HTTP 5xx (赤) / WARN = HTTP 4xx (黄)
# ==========================================================================
FROM='2026-08-20 09:00:00'                      # 検索開始 JST (日付のみ / 秒省略も可)
TO='2026-08-20 18:00:00'                        # 検索終了 JST (空にすると現在まで)
CONTAINER='front|back'                          # 対象コンテナ    (例: back)
SERVICE='intra-web|intra-api|inter-api|sf-api'  # 対象ECSサービス (例: sf-api)
LEVEL=''                                        # レベル絞込      (空=全て / 例: ERROR|WARN)
STATUS=''                                       # ステータス絞込  (空=全て / 例: 404|5[0-9][0-9])
KEYWORD=''                                      # キーワード正規表現 (空=なし / 一致箇所を反転表示)
EXCLUDE=''                                      # 除外正規表現    (空=なし / 例: /health|\.(css|js|png))

TZ=JST-9 find /mnt/logs -mindepth 6 -maxdepth 6 -type f \
    -regextype posix-extended \
    -regex ".*/($CONTAINER)/logs/($SERVICE)/mid/[^/]+/access_log(\.[0-9]{4}-[0-9]{2}-[0-9]{2})?\.log" \
    -newermt "${FROM:-1970-01-01}" -print0 \
| KW="$KEYWORD" EX="$EXCLUDE" xargs -0 -r gawk \
    -v from="$FROM" -v to="$TO" -v lv="$LEVEL" -v st="$STATUS" '
    function norm(s, d) { return s substr(d, length(s) + 1) }
    function jst(m,   e) {
        e = mktime(m[3] " " MON[m[2]] " " m[1] " " m[4] " " m[5] " 00", 1)
        e -= (m[7] == "-" ? -1 : 1) * (m[8] * 3600 + m[9] * 60)
        return strftime("%Y-%m-%d %H:%M", e + 9 * 3600, 1)
    }
    BEGIN {
        f = norm(from, "0000-01-01 00:00:00")
        t = norm(to,   "9999-12-31 23:59:59")
        # 処理の長いリクエストは受付時刻で後から出力されるため、TO + 10 分を超えたら次のファイルへ
        stop = (to == "") ? t : strftime("%Y-%m-%d %H:%M:%S", mktime(gensub(/[-:]/, " ", "g", t), 1) + 600, 1)
        split("Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec", a, " ")
        for (i = 1; i <= 12; i++) MON[a[i]] = i
        kw = ENVIRON["KW"]; ex = ENVIRON["EX"]
        if (length(lv)) lre = "^(" toupper(lv) ")$"
        if (length(st)) sre = "^(" st ")$"
    }
    BEGINFILE {
        if (length(ERRNO)) { print "[SKIP] " FILENAME " : " ERRNO > "/dev/stderr"; nextfile }
        print "#F\037" FILENAME
        bad = 0
    }
    match($0, /\[([0-9]{2})\/([A-Z][a-z]{2})\/([0-9]{4}):([0-9]{2}):([0-9]{2}):([0-9]{2}) ([-+])([0-9]{2})([0-9]{2})\]/, m) {
        h = substr($0, RSTART + 1, 17) m[7] m[8] m[9]
        if ((h in J) == 0) J[h] = jst(m)
        key = J[h] ":" m[6]
        if (key > stop) nextfile
        if (key < f || key > t) next
        if (length(ex) && $0 ~ ex) next
        if (length(kw) && $0 !~ kw) next
        pre = substr($0, 1, RSTART - 1); msg = substr($0, RSTART + RLENGTH)
        code = match(msg, /"[^"]*" ([1-5][0-9][0-9])( |$)/, s) ? s[1] : "-"
        lvl = (code ~ /^5/) ? "ERROR" : (code ~ /^4/) ? "WARN" : "INFO"
        if (length(lre) && match(lvl, lre) == 0) next
        if (length(sre) && match(code, sre) == 0) next
        sub(/ +$/, "", pre); sub(/^ +/, "", msg); gsub(/\t/, " ", msg)
        print key "\037" FILENAME "\037" lvl "\037" code "\037" pre (pre != "" && msg != "" ? " " : "") msg
        next
    }
    NF { bad++ }
    ENDFILE { if (bad) print "#U\037" FILENAME "\037" bad }
' \
| LC_ALL=C sort -s -t $'\037' -k1,1 \
| KW="$KEYWORD" EX="$EXCLUDE" gawk \
    -v from="$FROM" -v to="$TO" -v lv="$LEVEL" -v st="$STATUS" \
    -v ctr="$CONTAINER" -v svc="$SERVICE" -v cols="${COLUMNS:-$(tput cols)}" '
    function norm(s, d) { return s substr(d, length(s) + 1) }
    function rep(c, n,   s) { s = sprintf("%" n "s", ""); gsub(/ /, c, s); return s }
    function hl(s) { if (length(kw)) gsub(kw, ESC "7m&" ESC "27m", s); return s }
    function num(n, c) { return (n > 0 ? c : "") sprintf("%7d", n) R }
    function opt(v, d) { return v == "" ? d : v }
    function meta(p,   a, n) {
        n = split(p, a, "/")
        C[p] = a[n-5]; S[p] = a[n-3]
        T[p] = length(a[n-1]) == 32 ? substr(a[n-1], 1, 8) : a[n-1]
    }
    BEGIN {
        FS = "\037"; ESC = "\033["; R = ESC "0m"; G = ESC "90m"; sep = G " | " R
        TAG["ERROR"] = ESC "97;41m"; MSG["ERROR"] = ESC "31m"; SC["5"] = ESC "1;31m"
        TAG["WARN"]  = ESC "30;43m"; MSG["WARN"]  = ESC "33m"; SC["4"] = ESC "1;33m"
        TAG["INFO"]  = ESC "32m";    SC["2"] = ESC "32m";      SC["1"] = SC["3"] = ESC "36m"; SC["-"] = G
        CC["front"]  = ESC "36m";    CC["back"]   = ESC "35m"
        kw = ENVIRON["KW"]; ex = ENVIRON["EX"]
        f = norm(from, "0000-01-01 00:00:00")
        t = norm(to,   "9999-12-31 23:59:59")
        w = (cols + 0 >= 40) ? cols - 1 : 79
        HR = rep("=", w); LR = rep("-", w)
        period = (from == "" ? "(指定なし)" : f) "  ～  " (to == "" ? "(現在まで)" : t)
        print HR
        print "  JBoss EAP access_log 検索"
        print "    期間 (JST) : " period
        print "    時刻変換   : ログ時刻のオフセット (例: +0000) を基に JST (UTC+9) に変換して表示・ソート"
        print "    対象       : コンテナ=" ctr "   サービス=" svc
        print "    絞り込み   : レベル=" opt(lv, "(全て)") "   ステータス=" opt(st, "(全て)")
        print "                 キーワード=" opt(kw, "(なし)") "   除外=" opt(ex, "(なし)")
        print "    色分け     : " TAG["ERROR"] " ERROR " R " HTTP 5xx    " TAG["WARN"] " WARN " R " HTTP 4xx"
        re = "^[0-9]{4}-[0-9]{2}-[0-9]{2} [0-9]{2}:[0-9]{2}:[0-9]{2}$"
        if (match(f, re) == 0 || match(t, re) == 0) {
            print "  " TAG["ERROR"] " FROM / TO の書式が不正です (YYYY-MM-DD HH:MM:SS) " R
            print HR; bad = 1; exit 1
        }
        print LR
        printf "%-19s%s%-5s%s%-9s%s%-8s%s%-5s%s%-3s%s%s\n", "TIMESTAMP (JST)", sep, "CTR", sep, "SERVICE", sep, "TASK", sep, "LEVEL", sep, "STS", sep, "ACCESS LOG"
        print LR
    }
    $1 == "#F" { if (($2 in C) == 0) meta($2); H[$2] += 0; GH[C[$2] " / " S[$2]] += 0; nfile++; next }
    $1 == "#U" { U[$2] = $3; tu += $3; next }
    {
        p = $2; if ((p in C) == 0) meta(p)
        lvl = $3; code = $4
        printf "%s%s%s%-5s%s%s%-9s%s%s%-8s%s%s%s%-5s%s%s%s%-3s%s%s%s%s%s\n",
            $1, sep, CC[C[p]], C[p], R, sep, S[p], sep, G, T[p], R, sep,
            TAG[lvl], lvl, R, sep, SC[substr(code, 1, 1)], code, R, sep, MSG[lvl], hl($5), R
        g = C[p] " / " S[p]; H[p]++; GH[g]++; CODE[code]++; total++
        if (lvl == "ERROR") { HE[p]++; GE[g]++; te++ }
        if (lvl == "WARN")  { HW[p]++; GW[g]++; tw++ }
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
        printf "  %-8s", "STATUS"; x = 10
        for (c in CODE) {
            if (x + length(c "=" CODE[c]) + 2 > w) { printf "\n%10s", ""; x = 10 }
            printf "%s%s%s=%d  ", SC[substr(c, 1, 1)], c, R, CODE[c]; x += length(c "=" CODE[c]) + 2
        }
        print (total ? "" : "(なし)")
        print LR
        printf "  %7s%7s%7s   %s\n", "HITS", "ERROR", "WARN", "FILE (" (nfile + 0) " 件)"
        for (p in H) printf "  %7d%s%s   %s%s\n", H[p], num(HE[p], MSG["ERROR"]), num(HW[p], MSG["WARN"]), p,
            ((p in U) ? "  " TAG["WARN"] " 時刻解析不可 " U[p] " 行 " R : "")
        print LR
        printf "  %-24s%7d%s%s\n", "TOTAL", total, num(te, MSG["ERROR"]), num(tw, MSG["WARN"])
        if (tu) print "  " TAG["WARN"] " 時刻を解析できない行が " tu " 行あります (access-log の pattern を確認してください) " R
        print HR
    }
'
