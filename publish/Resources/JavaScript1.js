var e;
function ke(tfxh) {
    var arrxh = new Array();
    arrxh[0] = tfxh;
    e = {
        appid: "201010",
        appkey: "72933362EAA649B893699E6191BC898F",
        token: "",
        system: undefined,
        timestamp: undefined,
        rank: "",
        version: "1.0.0",
        param: {
            input: arrxh,
            VendorIds: []
        }
    }
    var t = e.appid,
        r = e.appkey,
        n = e.token,
        o = void 0 === n ? "" : n,
        i = e.version,
        a = void 0 === i ? "1.0.0" : i,
        s = e.param,
        c = e.system,
        u = e.timestamp,
        l = e.rank;
    u = u || Se.call(Ee(), "yyyy/MM/dd hh:mm:ss"),
        l = l || Ce(16),
        c = "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.198 Safari/537.36";
    var f = {
        system: c,
        version: a
    }, d = {
        apiVersion: "1.0"
    }, h = {
        appid: t,
        client: f,
        param: s,
        timestamp: u,
        rank: l,
        sign: ""
    }, p = "appid=" + t + "&client=" + JSON.stringify(f) + "&param=" + JSON.stringify(s) + "&rank=" + l + "&timestamp=" + u + "&token=" + o + "&key=" + r;

    h.sign = hash(p);
    var shuchu = h.sign + "-" + h.rank + "-" + h.timestamp;
    return shuchu
}


function Ee() {
    var e = 480,
        t = (new Date).getTimezoneOffset(),
        r = 60 * (e + t) * 1e3,
        n = (new Date).getTime(),
        o = new Date(n + r);
    return o
}

function Se(e) {
    var t = {
        "M+": this.getMonth() + 1,
        "d+": this.getDate(),
        "h+": this.getHours(),
        "m+": this.getMinutes(),
        "s+": this.getSeconds(),
        "q+": Math.floor((this.getMonth() + 3) / 3),
        S: this.getMilliseconds()
    };
    for (var r in /(y+)/.test(e) && (e = e.replace(RegExp.$1, (this.getFullYear() + "").substr(4 - RegExp.$1.length))),
        t)
        new RegExp("(" + r + ")").test(e) && (e = e.replace(RegExp.$1, 1 === RegExp.$1.length ? t[r] : ("00" + t[r]).substr(("" + t[r]).length)));
    return e
}

function Ce(e) {
    var t = Math.random().toString().slice(2);
    if (t.length === e) return t;
    if (t.length > e) return t.slice(0, e);
    for (var r = e - t.length, n = 0; n < r; n++)
        t += "0";
    return t
}

function p() {
    this.reset()
}

hash = function (e, t) {
    e.reset;
    return hashBinary(u(e), t)
}
hashBinary = function (e, t) {
    var r = i(e),
        n = c(r);
    return t ? h(n) : n
}

function u(e) {
    return /[\u0080-\uFFFF]/.test(e) && (e = unescape(encodeURIComponent(e))),
        e
}

function i(e) {
    var t, o, i, a, s, c, u = e.length,
        l = [1732584193, -271733879, -1732584194, 271733878];
    for (t = 64; t <= u; t += 64)
        r(l, n(e.substring(t - 64, t)));
    for (e = e.substring(t - 64),
        o = e.length,
        i = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
        t = 0; t < o; t += 1)
        i[t >> 2] |= e.charCodeAt(t) << (t % 4 << 3);
    if (i[t >> 2] |= 128 << (t % 4 << 3),
        t > 55) for (r(l, i),
            t = 0; t < 16; t += 1)
            i[t] = 0;
    return a = 8 * u,
        a = a.toString(16).match(/(.*?)(.{0,8})$/),
        s = parseInt(a[2], 16),
        c = parseInt(a[1], 16) || 0,
        i[14] = s,
        i[15] = c,
        r(l, i),
        l
}

function n(e) {
    var t, r = [];
    for (t = 0; t < 64; t += 4)
        r[t >> 2] = e.charCodeAt(t) + (e.charCodeAt(t + 1) << 8) + (e.charCodeAt(t + 2) << 16) + (e.charCodeAt(t + 3) << 24);
    return r
}

function r(e, t) {
    var r = e[0],
        n = e[1],
        o = e[2],
        i = e[3];
    r += (n & o | ~n & i) + t[0] - 680876936 | 0,
        r = (r << 7 | r >>> 25) + n | 0,
        i += (r & n | ~r & o) + t[1] - 389564586 | 0,
        i = (i << 12 | i >>> 20) + r | 0,
        o += (i & r | ~i & n) + t[2] + 606105819 | 0,
        o = (o << 17 | o >>> 15) + i | 0,
        n += (o & i | ~o & r) + t[3] - 1044525330 | 0,
        n = (n << 22 | n >>> 10) + o | 0,
        r += (n & o | ~n & i) + t[4] - 176418897 | 0,
        r = (r << 7 | r >>> 25) + n | 0,
        i += (r & n | ~r & o) + t[5] + 1200080426 | 0,
        i = (i << 12 | i >>> 20) + r | 0,
        o += (i & r | ~i & n) + t[6] - 1473231341 | 0,
        o = (o << 17 | o >>> 15) + i | 0,
        n += (o & i | ~o & r) + t[7] - 45705983 | 0,
        n = (n << 22 | n >>> 10) + o | 0,
        r += (n & o | ~n & i) + t[8] + 1770035416 | 0,
        r = (r << 7 | r >>> 25) + n | 0,
        i += (r & n | ~r & o) + t[9] - 1958414417 | 0,
        i = (i << 12 | i >>> 20) + r | 0,
        o += (i & r | ~i & n) + t[10] - 42063 | 0,
        o = (o << 17 | o >>> 15) + i | 0,
        n += (o & i | ~o & r) + t[11] - 1990404162 | 0,
        n = (n << 22 | n >>> 10) + o | 0,
        r += (n & o | ~n & i) + t[12] + 1804603682 | 0,
        r = (r << 7 | r >>> 25) + n | 0,
        i += (r & n | ~r & o) + t[13] - 40341101 | 0,
        i = (i << 12 | i >>> 20) + r | 0,
        o += (i & r | ~i & n) + t[14] - 1502002290 | 0,
        o = (o << 17 | o >>> 15) + i | 0,
        n += (o & i | ~o & r) + t[15] + 1236535329 | 0,
        n = (n << 22 | n >>> 10) + o | 0,
        r += (n & i | o & ~i) + t[1] - 165796510 | 0,
        r = (r << 5 | r >>> 27) + n | 0,
        i += (r & o | n & ~o) + t[6] - 1069501632 | 0,
        i = (i << 9 | i >>> 23) + r | 0,
        o += (i & n | r & ~n) + t[11] + 643717713 | 0,
        o = (o << 14 | o >>> 18) + i | 0,
        n += (o & r | i & ~r) + t[0] - 373897302 | 0,
        n = (n << 20 | n >>> 12) + o | 0,
        r += (n & i | o & ~i) + t[5] - 701558691 | 0,
        r = (r << 5 | r >>> 27) + n | 0,
        i += (r & o | n & ~o) + t[10] + 38016083 | 0,
        i = (i << 9 | i >>> 23) + r | 0,
        o += (i & n | r & ~n) + t[15] - 660478335 | 0,
        o = (o << 14 | o >>> 18) + i | 0,
        n += (o & r | i & ~r) + t[4] - 405537848 | 0,
        n = (n << 20 | n >>> 12) + o | 0,
        r += (n & i | o & ~i) + t[9] + 568446438 | 0,
        r = (r << 5 | r >>> 27) + n | 0,
        i += (r & o | n & ~o) + t[14] - 1019803690 | 0,
        i = (i << 9 | i >>> 23) + r | 0,
        o += (i & n | r & ~n) + t[3] - 187363961 | 0,
        o = (o << 14 | o >>> 18) + i | 0,
        n += (o & r | i & ~r) + t[8] + 1163531501 | 0,
        n = (n << 20 | n >>> 12) + o | 0,
        r += (n & i | o & ~i) + t[13] - 1444681467 | 0,
        r = (r << 5 | r >>> 27) + n | 0,
        i += (r & o | n & ~o) + t[2] - 51403784 | 0,
        i = (i << 9 | i >>> 23) + r | 0,
        o += (i & n | r & ~n) + t[7] + 1735328473 | 0,
        o = (o << 14 | o >>> 18) + i | 0,
        n += (o & r | i & ~r) + t[12] - 1926607734 | 0,
        n = (n << 20 | n >>> 12) + o | 0,
        r += (n ^ o ^ i) + t[5] - 378558 | 0,
        r = (r << 4 | r >>> 28) + n | 0,
        i += (r ^ n ^ o) + t[8] - 2022574463 | 0,
        i = (i << 11 | i >>> 21) + r | 0,
        o += (i ^ r ^ n) + t[11] + 1839030562 | 0,
        o = (o << 16 | o >>> 16) + i | 0,
        n += (o ^ i ^ r) + t[14] - 35309556 | 0,
        n = (n << 23 | n >>> 9) + o | 0,
        r += (n ^ o ^ i) + t[1] - 1530992060 | 0,
        r = (r << 4 | r >>> 28) + n | 0,
        i += (r ^ n ^ o) + t[4] + 1272893353 | 0,
        i = (i << 11 | i >>> 21) + r | 0,
        o += (i ^ r ^ n) + t[7] - 155497632 | 0,
        o = (o << 16 | o >>> 16) + i | 0,
        n += (o ^ i ^ r) + t[10] - 1094730640 | 0,
        n = (n << 23 | n >>> 9) + o | 0,
        r += (n ^ o ^ i) + t[13] + 681279174 | 0,
        r = (r << 4 | r >>> 28) + n | 0,
        i += (r ^ n ^ o) + t[0] - 358537222 | 0,
        i = (i << 11 | i >>> 21) + r | 0,
        o += (i ^ r ^ n) + t[3] - 722521979 | 0,
        o = (o << 16 | o >>> 16) + i | 0,
        n += (o ^ i ^ r) + t[6] + 76029189 | 0,
        n = (n << 23 | n >>> 9) + o | 0,
        r += (n ^ o ^ i) + t[9] - 640364487 | 0,
        r = (r << 4 | r >>> 28) + n | 0,
        i += (r ^ n ^ o) + t[12] - 421815835 | 0,
        i = (i << 11 | i >>> 21) + r | 0,
        o += (i ^ r ^ n) + t[15] + 530742520 | 0,
        o = (o << 16 | o >>> 16) + i | 0,
        n += (o ^ i ^ r) + t[2] - 995338651 | 0,
        n = (n << 23 | n >>> 9) + o | 0,
        r += (o ^ (n | ~i)) + t[0] - 198630844 | 0,
        r = (r << 6 | r >>> 26) + n | 0,
        i += (n ^ (r | ~o)) + t[7] + 1126891415 | 0,
        i = (i << 10 | i >>> 22) + r | 0,
        o += (r ^ (i | ~n)) + t[14] - 1416354905 | 0,
        o = (o << 15 | o >>> 17) + i | 0,
        n += (i ^ (o | ~r)) + t[5] - 57434055 | 0,
        n = (n << 21 | n >>> 11) + o | 0,
        r += (o ^ (n | ~i)) + t[12] + 1700485571 | 0,
        r = (r << 6 | r >>> 26) + n | 0,
        i += (n ^ (r | ~o)) + t[3] - 1894986606 | 0,
        i = (i << 10 | i >>> 22) + r | 0,
        o += (r ^ (i | ~n)) + t[10] - 1051523 | 0,
        o = (o << 15 | o >>> 17) + i | 0,
        n += (i ^ (o | ~r)) + t[1] - 2054922799 | 0,
        n = (n << 21 | n >>> 11) + o | 0,
        r += (o ^ (n | ~i)) + t[8] + 1873313359 | 0,
        r = (r << 6 | r >>> 26) + n | 0,
        i += (n ^ (r | ~o)) + t[15] - 30611744 | 0,
        i = (i << 10 | i >>> 22) + r | 0,
        o += (r ^ (i | ~n)) + t[6] - 1560198380 | 0,
        o = (o << 15 | o >>> 17) + i | 0,
        n += (i ^ (o | ~r)) + t[13] + 1309151649 | 0,
        n = (n << 21 | n >>> 11) + o | 0,
        r += (o ^ (n | ~i)) + t[4] - 145523070 | 0,
        r = (r << 6 | r >>> 26) + n | 0,
        i += (n ^ (r | ~o)) + t[11] - 1120210379 | 0,
        i = (i << 10 | i >>> 22) + r | 0,
        o += (r ^ (i | ~n)) + t[2] + 718787259 | 0,
        o = (o << 15 | o >>> 17) + i | 0,
        n += (i ^ (o | ~r)) + t[9] - 343485551 | 0,
        n = (n << 21 | n >>> 11) + o | 0,
        e[0] = r + e[0] | 0,
        e[1] = n + e[1] | 0,
        e[2] = o + e[2] | 0,
        e[3] = i + e[3] | 0
}
function c(e) {
    var t;
    for (t = 0; t < e.length; t += 1)
        e[t] = s(e[t]);
    return e.join("")
}
function s(e) {
    var r, n = "";
    var t = ["0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "a", "b", "c", "d", "e", "f"];
    for (r = 0; r < 4; r += 1)
        n += t[e >> 8 * r + 4 & 15] + t[e >> 8 * r & 15];
    return n
}