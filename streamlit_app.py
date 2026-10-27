# -*- coding: utf-8 -*-
"""
Asisten Cleaning Data SKP 2026
==============================
Aplikasi Streamlit untuk membantu Tim SKP BPS Kab/Kota melakukan cleaning data
hasil pendataan lapangan SKP 2026 sesuai dokumen "Mekanisme Cleaning Data SKP 2026".

Fitur:
  1. Upload data (CSV / Excel) hasil unduhan FASIH SM + pemetaan kolom otomatis.
  2. Deteksi 21 anomali (A_01 - A_21) dan 14 outlier (O_01 - O_14).
  3. Dashboard prioritas RTUP (assignment) yang perlu di-cleaning.
  4. Review per RTUP: edit data, bandingkan status flag, isi justifikasi.
  5. Pembuatan teks "Blok Catatan Anomali" siap salin ke FASIH SM
     (format  #[KODE]: Diperbaiki / Sesuai Kondisi Lapangan / Non Respon# penjelasan).
  6. Flagging di luar daftar (B_xx, BPS Provinsi).
  7. Unduh hasil (Excel / CSV), log perubahan, dan data bersih.

Jalankan:  streamlit run app.py
"""
import io
import json
import re
from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Asisten Cleaning Data SKP 2026",
    page_icon="🌾",
    layout="wide",
)

# =============================================================================
# 1. KONSTANTA (sesuai pedoman)
# =============================================================================
ANOMALI_DESC = {
    "A_01": "Nama KRT (R110) kurang dari 3 huruf",
    "A_02": "Nama KRT (R110) dengan tiga huruf yang sama secara berurutan",
    "A_03": 'NIK Kepala Rumah Tangga (R111) terisi " - ". Dapat terisi strip (-), jika KRT tidak mempunyai NIK.',
    "A_04": "NIK KRT (R111) dengan lima angka yang sama secara berurutan",
    "A_05": "Nomor Telepon/HP (R112) dengan lima angka yang sama secara berurutan (misal: 081261399999)",
    "A_06": 'Kepala Rumah Tangga (R304="1") tetapi berumur kurang dari 10 tahun (R306<10)',
    "A_07": "Jam kerja (R326.h) dalam seminggu melebihi 112 jam",
    "A_08": "Upah pekerja di luar Sektor Pertanian (R329) sebesar Rp 1.000.000 ke bawah",
    "A_09": "Umur 10 tahun keatas tetapi belum pernah bersekolah",
    "A_10": "Umur 50 tahun keatas tetapi masih bersekolah",
    "A_11": "Nomor ID Pelanggan PLN (R409c.ii) dengan lima angka yang sama secara berurutan.",
    "A_12": "Nomor meteran Pelanggan PLN (R409c.iii) dengan lima angka yang sama secara berurutan.",
    "A_13": "Luas lantai bangunan tempat tinggal (R402) dibagi jumlah ART (R301) kurang dari 3,6 meter persegi",
    "A_14": "Bahan atap terluas (R403.a) Lainnya, bahan dinding terluas (R404.a) Lainnya, dan Bahan lantai terluas (R405.a) Lainnya.",
    "A_15": "Atap terluas (R403.a) terisi beton tetapi dinding terluas (R404.a) selain tembok",
    "A_16": "Jenis kloset (R406b) berupa leher angsa dengan tempat pembuangan akhir tinja (R406c) ke pantai/tanah lapang/kebun atau lainnya",
    "A_17": "Sumber air utama untuk mandi/cuci/dll (R408a) berupa air kemasan atau air isi ulang",
    "A_18": "Listrik non PLN tetapi pengeluaran kurang dari 25 ribu",
    "A_19": 'Pendapatan Usaha Pertanian (R608k5) bernilai negatif atau nol atau lebih besar dari Rp 1.000.000.000,- (1 milyar rupiah). Berikan catatan mengapa petani mengalami kerugian. Pastikan kembali satuan dalam "ribu rupiah".',
    "A_20": "Luas Lahan Pertanian yang Dikuasai (R701b.iv.k2) terisi 0. Luas Lahan Pertanian yang Dikuasai dapat bernilai 0, jika usaha perikanan tangkap, usaha peternakan dengan sistem dilepas, usaha penangkapan satwa liar, usaha pemungutan hasil hutan.",
    "A_21": "Subsektor yang diusahakan hanya perikanan tangkap tetapi luas lahan yang dikuasai lebih dari 0",
}
OUTLIER_DESC = {
    "O_01": "Luas lahan yang dikuasai RTUP",
    "O_02": "Jumlah ternak (ekor) yang dikuasai RTUP",
    "O_03": "Pendapatan usaha tanaman pangan",
    "O_04": "Pendapatan usaha tanaman hortikultura",
    "O_05": "Pendapatan usaha tanaman perkebunan",
    "O_06": "Pendapatan usaha peternakan",
    "O_07": "Pendapatan usaha perikanan tangkap",
    "O_08": "Pendapatan usaha perikanan budidaya",
    "O_09": "Pendapatan usaha kehutanan",
    "O_10": "Pendapatan usaha jasa pertanian",
    "O_11": "Jumlah Pendapatan usaha pertanian",
    "O_12": "Jumlah Pendapatan di luar usaha pertanian",
    "O_13": "Jumlah Pendapatan (608_k5 + 620)",
    "O_14": "Share Pendapatan usaha pertanian (608_k5 / (608_k5 + 620) * 100)",
}
FLAG_DESC = {**ANOMALI_DESC, **OUTLIER_DESC}

# (kode outlier, kunci kolom kanonik). "*sum" & "*share" dihitung otomatis.
OUTLIERS = [
    ("O_01", "o701b"), ("O_02", "o702b"), ("O_03", "o601"), ("O_04", "o602"),
    ("O_05", "o603"), ("O_06", "o604"), ("O_07", "o605a"), ("O_08", "o605b"),
    ("O_09", "o606"), ("O_10", "o607"), ("O_11", "r608k5"), ("O_12", "o620"),
    ("O_13", "*sum"), ("O_14", "*share"),
]

STATUS_PLACEHOLDER = "— belum dipilih —"
STATUS_OPTS = [STATUS_PLACEHOLDER, "Diperbaiki", "Sesuai Kondisi Lapangan", "Non Respon"]


def _al(c):
    return [c, "r" + c, c + "k5", "r" + c + "k5"]


# key, label, grup, alias nama kolom (setelah dinormalisasi: huruf kecil tanpa simbol)
FIELDS = [
    ("id", "ID RTUP / Assignment  (WAJIB)", "Identitas", ["idassignment", "assignmentid", "assignment", "idrtup", "rtup", "kodeassignment", "kodert"]),
    ("status", "Status assignment (opsional)", "Identitas", ["statusassignment", "assignmentstatus", "statusdata"]),
    ("grp", "Pengelompokan outlier, mis. kecamatan/desa (opsional)", "Identitas", ["kecamatan", "desa"]),
    ("r110", "R110 – Nama KRT", "Blok 1", []),
    ("r111", "R111 – NIK KRT", "Blok 1", ["nik", "nikkrt"]),
    ("r112", "R112 – Nomor Telepon/HP", "Blok 1", ["nohp", "telepon"]),
    ("r301", "R301 – Jumlah ART", "Blok 3", ["jumlahart"]),
    ("r304", "R304 – Hubungan dengan KRT (1 = KRT)", "Blok 3", []),
    ("r306", "R306 – Umur", "Blok 3", ["umur"]),
    ("r314", "R314 – Partisipasi sekolah (1 = belum pernah, 2 = masih sekolah)", "Blok 3", []),
    ("r326h", "R326.h – Jam kerja seminggu", "Blok 3", ["jamkerja"]),
    ("r329", "R329 – Upah pekerja di luar sektor pertanian", "Blok 3", []),
    ("r402", "R402 – Luas lantai (m²)", "Blok 4", ["luaslantai"]),
    ("r403a", "R403.a – Bahan atap terluas", "Blok 4", []),
    ("r404a", "R404.a – Bahan dinding terluas", "Blok 4", []),
    ("r405a", "R405.a – Bahan lantai terluas", "Blok 4", []),
    ("r406b", "R406b – Jenis kloset", "Blok 4", []),
    ("r406c", "R406c – Tempat pembuangan akhir tinja", "Blok 4", []),
    ("r408a", "R408a – Sumber air utama mandi/cuci", "Blok 4", []),
    ("r409a", "R409a – Sumber penerangan (kode 3 = non PLN)", "Blok 4", []),
    ("r409cii", "R409c.ii – No. ID Pelanggan PLN", "Blok 4", []),
    ("r409ciii", "R409c.iii – No. Meteran PLN", "Blok 4", []),
    ("r411", "R411 – Pengeluaran listrik (ribu rupiah)", "Blok 4", []),
    ("r608k5", "R608.k5 – Jumlah pendapatan usaha pertanian (ribu rupiah)", "Blok 6", []),
    ("o601", "R601.k5 – Pendapatan tanaman pangan", "Blok 6", _al("601")),
    ("o602", "R602.k5 – Pendapatan hortikultura", "Blok 6", _al("602")),
    ("o603", "R603.k5 – Pendapatan perkebunan", "Blok 6", _al("603")),
    ("o604", "R604.k5 – Pendapatan peternakan", "Blok 6", _al("604")),
    ("o605a", "R605a.k5 – Pendapatan perikanan tangkap", "Blok 6", _al("605a")),
    ("o605b", "R605b.k5 – Pendapatan perikanan budidaya", "Blok 6", _al("605b")),
    ("o606", "R606.k5 – Pendapatan kehutanan", "Blok 6", _al("606")),
    ("o607", "R607.k5 – Pendapatan jasa pertanian", "Blok 6", _al("607")),
    ("o620", "R620 – Pendapatan di luar usaha pertanian", "Blok 6", ["620", "r620"]),
    ("r120a", "R120a – Subsektor tanaman pangan (1 = ya, 2 = tidak)", "Blok 7", []),
    ("r120b", "R120b – Subsektor hortikultura", "Blok 7", []),
    ("r120c", "R120c – Subsektor perkebunan", "Blok 7", []),
    ("r120d", "R120d – Subsektor peternakan", "Blok 7", []),
    ("r120e", "R120e – Subsektor perikanan tangkap", "Blok 7", []),
    ("r120f", "R120f – Subsektor perikanan budidaya", "Blok 7", []),
    ("r120g", "R120g – Subsektor kehutanan", "Blok 7", []),
    ("r701bivk2", "R701b.iv.k2 – Luas lahan pertanian yang dikuasai", "Blok 7", ["r701bk2"]),
    ("o701b", "R701b – Luas lahan dikuasai RTUP (untuk outlier O_01)", "Blok 7", ["701b", "r701b", "r701bivk2"]),
    ("o702b", "R702b – Jumlah ternak (ekor) (untuk outlier O_02)", "Blok 7", ["702b", "r702b"]),
]
FIELD_LABEL = {f[0]: f[1] for f in FIELDS}

DEFAULT_P = dict(
    jam_max=112, upah_max=1_000_000, rasio_lantai=3.6, listrik_max=25,
    tani_max=1_000_000, method="IQR", k_iqr=3.0, k_z=3.5, ignore_zero=True,
)


# =============================================================================
# 2. FUNGSI BANTU
# =============================================================================
def norm(s) -> str:
    return re.sub(r"[^a-z0-9]", "", str(s).lower())


def to_num_value(x):
    if x is None:
        return np.nan
    x = str(x).strip()
    if x == "" or x.lower() in ("nan", "none", "-"):
        return np.nan
    x = x.replace("Rp", "").replace("rp", "").replace(" ", "")
    if re.fullmatch(r"-?\d{1,3}(\.\d{3})+(,\d+)?", x):  # format Indonesia 4.000.000,50
        x = x.replace(".", "").replace(",", ".")
    else:
        x = x.replace(",", ".")
    try:
        return float(x)
    except ValueError:
        return np.nan


RE_DIGIT5 = re.compile(r"(\d)\1{4}")
RE_HURUF3 = re.compile(r"([A-Za-z])\1\1", re.I)


def has_repeat5(v: str) -> bool:
    return bool(RE_DIGIT5.search(re.sub(r"[\s\-]", "", v)))


def fmt_num(v, dec=0):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return "-"
    s = f"{v:,.{dec}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


class Ctx:
    """Penyedia kolom kanonik (teks, angka, kode) dengan cache."""

    def __init__(self, canon: pd.DataFrame):
        self.c = canon
        self._n, self._k, self._ks = {}, {}, {}

    def s(self, k):
        return self.c[k].fillna("").astype(str).str.strip()

    def n(self, k):
        if k not in self._n:
            self._n[k] = self.c[k].map(to_num_value).astype(float)
        return self._n[k]

    def k(self, k):
        if k not in self._k:
            ext = self.s(k).str.extract(r"^(\d+)")[0]
            self._k[k] = ext.map(lambda v: str(int(v)) if isinstance(v, str) else None)
        return self._k[k]

    def ks(self, k):
        if k not in self._ks:
            def conv(v):
                v = re.sub(r"\.0+\b", "", str(v))
                return frozenset(re.findall(r"\d+", v))
            self._ks[k] = self.s(k).map(conv)
        return self._ks[k]


# =============================================================================
# 3. ATURAN ANOMALI (A_01 s.d. A_21)
# =============================================================================
def _r120_lain(x, huruf, nilai):
    m = pd.Series(True, index=x.c.index)
    for h in huruf:
        m &= x.k("r120" + h) == nilai
    return m


RULES = [
    dict(kode="A_01", req=["r110"], fn=lambda x, P: x.s("r110").str.len() < 3),
    dict(kode="A_02", req=["r110"], fn=lambda x, P: x.s("r110").map(lambda v: bool(RE_HURUF3.search(v)))),
    dict(kode="A_03", req=["r111"], fn=lambda x, P: x.s("r111") == "-"),
    dict(kode="A_04", req=["r111"], fn=lambda x, P: x.s("r111").map(has_repeat5)),
    dict(kode="A_05", req=["r112"], fn=lambda x, P: x.s("r112").map(has_repeat5)),
    dict(kode="A_06", req=["r304", "r306"], fn=lambda x, P: (x.k("r304") == "1") & (x.n("r306") < 10)),
    dict(kode="A_07", req=["r326h"], fn=lambda x, P: x.n("r326h") > P["jam_max"]),
    dict(kode="A_08", req=["r329"], fn=lambda x, P: (x.n("r329") > 0) & (x.n("r329") <= P["upah_max"])),
    dict(kode="A_09", req=["r306", "r314"], fn=lambda x, P: (x.n("r306") >= 10) & (x.k("r314") == "1")),
    dict(kode="A_10", req=["r306", "r314"], fn=lambda x, P: (x.n("r306") >= 50) & (x.k("r314") == "2")),
    dict(kode="A_11", req=["r409cii"], fn=lambda x, P: x.s("r409cii").map(has_repeat5)),
    dict(kode="A_12", req=["r409ciii"], fn=lambda x, P: x.s("r409ciii").map(has_repeat5)),
    dict(kode="A_13", req=["r402", "r301"],
         fn=lambda x, P: (x.n("r402") / x.n("r301").where(x.n("r301") > 0)) < P["rasio_lantai"]),
    dict(kode="A_14", req=["r403a", "r404a", "r405a"],
         fn=lambda x, P: (x.k("r403a") == "8") & (x.k("r404a") == "7") & (x.k("r405a") == "9")),
    dict(kode="A_15", req=["r403a", "r404a"],
         fn=lambda x, P: (x.k("r403a") == "1") & x.k("r404a").notna() & (x.k("r404a") != "1")),
    dict(kode="A_16", req=["r406b", "r406c"],
         fn=lambda x, P: (x.k("r406b") == "1") & x.k("r406c").isin(["5", "6"])),
    dict(kode="A_17", req=["r408a"], fn=lambda x, P: x.k("r408a").isin(["1", "2"])),
    dict(kode="A_18", req=["r409a", "r411"],
         fn=lambda x, P: x.ks("r409a").map(lambda s: s == frozenset({"3"})) & (x.n("r411") <= P["listrik_max"])),
    dict(kode="A_19", req=["r608k5"], fn=lambda x, P: (x.n("r608k5") <= 0) | (x.n("r608k5") > P["tani_max"])),
    dict(kode="A_20", req=["r701bivk2", "r120a", "r120b", "r120c", "r120d", "r120e", "r120f", "r120g"],
         fn=lambda x, P: (x.n("r701bivk2") == 0) & _r120_lain(x, "abc", "2")
         & ((x.k("r120d") == "1") | (x.k("r120e") == "1") | (x.k("r120f") == "1") | (x.k("r120g") == "1"))),
    dict(kode="A_21", req=["r701bivk2", "r120a", "r120b", "r120c", "r120d", "r120e", "r120f", "r120g"],
         fn=lambda x, P: (x.k("r120e") == "1") & _r120_lain(x, "abcdfg", "2") & (x.n("r701bivk2") > 0)),
]

FLAG_FIELDS = {r["kode"]: list(r["req"]) for r in RULES}
for _kode, _fk in OUTLIERS:
    FLAG_FIELDS[_kode] = ["r608k5", "o620"] if _fk.startswith("*") else [_fk]


# =============================================================================
# 4. DETEKSI OUTLIER
# =============================================================================
def outlier_bounds(ref: pd.Series, method: str, k: float):
    ref = ref.dropna()
    if len(ref) < 8:
        return None
    if method == "IQR":
        q1, q3 = ref.quantile(0.25), ref.quantile(0.75)
        i = q3 - q1
        return q1 - k * i, q3 + k * i
    med = ref.median()
    mad = (ref - med).abs().median()
    if mad == 0:
        return None
    d = k * mad / 0.6745
    return med - d, med + d


def detect_outliers(canon, x, P, groups):
    recs = []
    fields = [fk for _, fk in OUTLIERS if not fk.startswith("*")]
    cols = {fk: x.n(fk) for fk in fields if fk in canon.columns}
    if not cols:
        return recs
    rt = pd.DataFrame(cols)
    rt["id"] = canon["id"]
    rt = rt.groupby("id", sort=False).first()
    if "r608k5" in rt and "o620" in rt:
        tot = rt[["r608k5", "o620"]].sum(axis=1, min_count=1)
        rt["*sum"] = tot
        rt["*share"] = (rt["r608k5"] / tot.where(tot > 0)) * 100
    g = canon.groupby("id", sort=False)["grp"].first() if "grp" in canon.columns else None
    k = P["k_iqr"] if P["method"] == "IQR" else P["k_z"]
    for kode, fk in OUTLIERS:
        if fk not in rt:
            continue
        v = rt[fk]
        dec = 1 if fk == "*share" else 0

        def ref_of(s):
            return s[s > 0] if P["ignore_zero"] else s

        glob = outlier_bounds(ref_of(v), P["method"], k)
        lo = pd.Series(np.nan, index=v.index)
        hi = pd.Series(np.nan, index=v.index)
        if g is not None:
            for _, ids in g.groupby(g).groups.items():
                b = outlier_bounds(ref_of(v.loc[ids]), P["method"], k) or glob
                if b:
                    lo.loc[ids], hi.loc[ids] = b
        elif glob:
            lo[:], hi[:] = glob
        mask = (v < lo) | (v > hi)
        if P["ignore_zero"]:
            mask &= v > 0
        for key in v.index[mask.fillna(False)]:
            info = f"nilai {fmt_num(v[key], dec)}; batas wajar {fmt_num(lo[key], dec)} s.d. {fmt_num(hi[key], dec)}"
            idx = tuple(int(i) for i in groups.get(key, []))
            recs.append((key, kode, "Outlier", info, idx))
    return recs


# =============================================================================
# 5. MESIN DETEKSI (cache)
# =============================================================================
@st.cache_data(show_spinner="Menjalankan pengecekan anomali & outlier …")
def compute_flags(df: pd.DataFrame, mapping: tuple, params_json: str):
    P = json.loads(params_json)
    m = {k: v for k, v in mapping if v}
    cols = {k: df[c] for k, c in m.items() if c in df.columns}
    empty = pd.DataFrame(columns=["id", "kode", "jenis", "deskripsi", "info", "rows", "n_baris"])
    if "id" not in cols or df.empty:
        return empty, []
    canon = pd.DataFrame(cols, index=df.index)
    canon["id"] = canon["id"].fillna("").astype(str).str.strip()
    canon = canon[canon["id"] != ""]
    if canon.empty:
        return empty, []
    x = Ctx(canon)
    groups = canon.groupby("id", sort=False).groups
    recs, skipped = [], []
    for r in RULES:
        miss = [f for f in r["req"] if f not in canon.columns]
        if miss:
            skipped.append((r["kode"], "kolom belum dipetakan: " + ", ".join(miss)))
            continue
        try:
            mask = r["fn"](x, P).fillna(False).astype(bool)
        except Exception as e:  # noqa
            skipped.append((r["kode"], f"error: {e}"))
            continue
        if mask.any():
            sub = canon.loc[mask, "id"]
            for key, idx in sub.index.groupby(sub.values).items():
                recs.append((key, r["kode"], "Anomali", "", tuple(int(i) for i in idx)))
    try:
        recs += detect_outliers(canon, x, P, groups)
    except Exception as e:  # noqa
        skipped.append(("O_xx", f"error: {e}"))
    if not recs:
        return empty, skipped
    out = pd.DataFrame(recs, columns=["id", "kode", "jenis", "info", "rows"])
    out["deskripsi"] = out["kode"].map(FLAG_DESC)
    out["n_baris"] = out["rows"].map(len)
    out = out.sort_values(["id", "kode"]).reset_index(drop=True)
    return out[["id", "kode", "jenis", "deskripsi", "info", "rows", "n_baris"]], skipped


def eligible_mask(df, m, only):
    if not only or not m.get("status"):
        return pd.Series(True, index=df.index)
    s = df[m["status"]].map(norm)
    return s.str.contains("approvedbypengawas|revokedbyadmin", regex=True)


# =============================================================================
# 6. DATA CONTOH & TEMPLATE
# =============================================================================
def make_demo(n=40, seed=7) -> pd.DataFrame:
    rng = np.random.default_rng(seed)
    nama = ["Slamet", "Budi Santoso", "Siti Aminah", "Joko", "Sumarni", "Bambang", "Rina Wati", "Dedi", "Ahmad Fauzi", "Sari"]
    digits = lambda k: "".join(map(str, rng.integers(0, 10, k)))  # noqa: E731
    rows = []
    for i in range(n):
        nart = int(rng.integers(1, 5))
        inc = {c: 0 for c in ["o601", "o602", "o603", "o604", "o605a", "o605b", "o606", "o607"]}
        inc["o601"] = int(rng.integers(5000, 40000))
        inc["o603"] = int(rng.integers(0, 20000))
        r608 = sum(inc.values())
        rt = dict(
            id=f"RTUP{i + 1:04d}",
            status="APPROVED BY Pengawas" if i % 19 != 18 else "SUBMITTED BY Pencacah",
            kecamatan=str(rng.choice(["Kec A", "Kec B"])),
            r110=str(rng.choice(nama)), r111=digits(16), r112="08" + digits(10),
            r301=nart, r402=int(rng.integers(30, 120)),
            r403a=2, r404a=1, r405a=int(rng.integers(3, 6)), r406b=1, r406c=1,
            r408a=int(rng.choice([3, 4, 5])), r409a=1,
            r409cii=digits(12), r409ciii=digits(11), r411=int(rng.integers(50, 300)),
            r608k5=r608, o620=int(rng.integers(10000, 60000)),
            r120a=1, r120b=2, r120c=2, r120d=2, r120e=2, r120f=2, r120g=2,
            r701bivk2=int(rng.integers(2000, 20000)), o702b=int(rng.integers(0, 6)),
            **inc,
        )
        rt["o701b"] = rt["r701bivk2"]
        # ---- suntik anomali / outlier ----
        if i % 8 == 1: rt["r110"] = "Aa"
        if i % 8 == 2: rt["r110"] = "Budiii"
        if i % 9 == 3: rt["r111"] = "-"
        if i % 10 == 4: rt["r112"] = "081211111999"
        if i % 11 == 5: rt["r402"] = nart * 2
        if i % 12 == 6: rt.update(r403a=8, r404a=7, r405a=9)
        if i % 13 == 7: rt.update(r403a=1, r404a=3)
        if i % 9 == 8: rt["r408a"] = 1
        if i % 10 == 9: rt.update(r409a="3", r411=10)
        if i % 7 == 0: rt["r608k5"] = 0
        if i % 14 == 10: rt["r608k5"] = 2_000_000
        if i % 15 == 11: rt.update(r120a=2, r120e=1, r701bivk2=5000)
        if i % 16 == 12: rt.update(r120a=2, r120d=1, r701bivk2=0)
        if i == 20: rt.update(o601=5_000_000, r608k5=5_000_000)
        if i == 25: rt["o620"] = 2_000_000
        for j in range(nart):
            umur = int(rng.integers(25, 60)) if j == 0 else int(rng.integers(1, 50))
            art = dict(
                r304=1 if j == 0 else int(rng.choice([2, 3, 4])), r306=umur,
                r314=3 if umur >= 10 else 2, r326h=int(rng.integers(20, 60)),
                r329=int(rng.integers(1_500_000, 4_000_000)) if j == 0 else 0,
            )
            if j == 0:
                if i % 11 == 2: art["r306"] = 8
                if i % 12 == 3: art["r326h"] = 120
                if i % 10 == 5: art["r329"] = 800_000
                if i % 13 == 4: art.update(r306=40, r314=1)
                if i % 17 == 5: art.update(r306=55, r314=2)
            rows.append({**rt, **art})
    return pd.DataFrame(rows).astype(str)


def template_csv() -> bytes:
    return (",".join(f[0] for f in FIELDS) + "\n").encode("utf-8")


# =============================================================================
# 7. STATE
# =============================================================================
ss = st.session_state
ss.setdefault("sig", None)
ss.setdefault("raw", None)
ss.setdefault("work", None)
ss.setdefault("just", {})
ss.setdefault("bfl", {})
ss.setdefault("log", [])
ss.setdefault("ver", 0)
ss.setdefault("flash", None)
ss.setdefault("demo", False)


def reset_state(sig, df):
    ss.sig = sig
    ss.raw = df.copy()
    ss.work = df.copy()
    ss.just, ss.bfl, ss.log, ss.ver, ss.flash = {}, {}, [], 0, None


@st.cache_data(show_spinner=False)
def read_upload(name: str, content: bytes, sheet=None):
    if name.lower().endswith(".csv"):
        df = pd.read_csv(io.BytesIO(content), dtype=str, keep_default_na=False, sep=None, engine="python")
    else:
        df = pd.read_excel(io.BytesIO(content), dtype=str, sheet_name=sheet or 0).fillna("")
    df.columns = [str(c).strip() for c in df.columns]
    return df.reset_index(drop=True)


def to_excel(sheets: dict) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for name, d in sheets.items():
            d.to_excel(w, sheet_name=name[:31], index=False)
    return buf.getvalue()


# =============================================================================
# 8. SIDEBAR
# =============================================================================
with st.sidebar:
    st.title("🌾 SKP 2026")
    st.caption("Asisten Cleaning Data · Tim SKP BPS Kab/Kota")
    up = st.file_uploader("Unggah data SKP (CSV / Excel)", type=["csv", "xlsx", "xls"])
    sheet = None
    if up is not None and not up.name.lower().endswith(".csv"):
        try:
            names = pd.ExcelFile(io.BytesIO(up.getvalue())).sheet_names
            if len(names) > 1:
                sheet = st.selectbox("Sheet", names)
        except Exception:
            pass
    c1, c2 = st.columns(2)
    if c1.button("Data contoh"):
        ss.demo = True
    c2.download_button("Template", template_csv(), "template_kolom_skp2026.csv", "text/csv")

    if up is not None:
        ss.demo = False
        try:
            sig = (up.name, up.size, sheet)
            if ss.sig != sig:
                reset_state(sig, read_upload(up.name, up.getvalue(), sheet))
        except Exception as e:
            st.error(f"Gagal membaca file: {e}")
    elif ss.demo and ss.sig != ("demo",):
        reset_state(("demo",), make_demo())

    st.divider()
    st.subheader("⚙️ Parameter deteksi")
    only_elig = st.checkbox(
        "Hanya status APPROVED BY Pengawas / REVOKED BY Admin", value=True,
        help="Sesuai pedoman, hanya data berstatus tersebut yang dapat di-cleaning. Berlaku jika kolom status dipetakan.",
    )
    with st.expander("Ambang batas anomali"):
        P = dict(DEFAULT_P)
        P["jam_max"] = st.number_input("A_07: jam kerja maks/minggu", value=112, step=1)
        P["upah_max"] = st.number_input(
            "A_08: batas upah (R329 ≤ nilai)", value=1_000_000, step=100_000,
            help="Pedoman menulis kondisi R329 <= 1000 namun deskripsi menyebut Rp 1.000.000. "
                 "Sesuaikan dengan satuan isian R329 pada data Anda.")
        P["rasio_lantai"] = st.number_input("A_13: luas lantai per ART (m²)", value=3.6, step=0.1)
        P["listrik_max"] = st.number_input("A_18: pengeluaran listrik (ribu rupiah)", value=25, step=1)
        P["tani_max"] = st.number_input("A_19: pendapatan tani maks (ribu rupiah)", value=1_000_000, step=100_000)
    with st.expander("Metode outlier"):
        P["method"] = st.radio("Metode", ["IQR", "Modified Z-score"], horizontal=True)
        if P["method"] == "IQR":
            P["k_iqr"] = st.slider("Pengali IQR", 1.0, 6.0, 3.0, 0.5)
        else:
            P["k_z"] = st.slider("Ambang |z|", 2.0, 8.0, 3.5, 0.5)
        P["ignore_zero"] = st.checkbox("Abaikan nilai 0 saat menghitung batas", value=True)
        st.caption("Pedoman merujuk dashboard FASIH untuk daftar outlier resmi; "
                   "metode ini hanya pendekatan pembantu bila dashboard belum tersedia.")

# =============================================================================
# 9. HALAMAN UTAMA
# =============================================================================
st.title("Asisten Cleaning Data SKP 2026")

if ss.raw is None:
    st.info(
        "👈 Unggah data hasil pendataan SKP dari FASIH SM (CSV/Excel) di sidebar, "
        "atau klik **Data contoh** untuk mencoba aplikasi.\n\n"
        "Format data: satu baris = satu RTUP, atau satu baris = satu ART dengan kolom "
        "**ID RTUP/assignment** yang sama berulang. Gunakan **Template** untuk nama kolom yang dikenali otomatis."
    )
    st.stop()

tab_map, tab_dash, tab_review, tab_out, tab_guide = st.tabs(
    ["1️⃣ Data & Kolom", "2️⃣ Dashboard", "3️⃣ Review & Cleaning", "4️⃣ Hasil & Unduh", "📖 Pedoman"]
)

# ---------------------------------------------------------------- TAB 1
with tab_map:
    raw = ss.raw
    st.write(f"**{len(raw):,}** baris · **{len(raw.columns)}** kolom")
    st.dataframe(raw.head(20))
    st.subheader("Pemetaan kolom")
    st.caption("Kolom dikenali otomatis dari nama. Periksa dan sesuaikan bila perlu. "
               "Aturan yang kolomnya belum dipetakan akan dilewati.")
    cols = list(raw.columns)
    ncol = {norm(c): c for c in cols}
    NONE = "(tidak ada)"
    mapping = {}
    groups_f = {}
    for key, label, grp, aliases in FIELDS:
        groups_f.setdefault(grp, []).append((key, label, aliases))
    sigk = str(ss.sig)
    for grp, items in groups_f.items():
        with st.expander(grp, expanded=(grp.startswith("Identitas"))):
            cc = st.columns(2)
            for n, (key, label, aliases) in enumerate(items):
                guess = next((ncol[a] for a in [key] + aliases if a in ncol), None)
                opts = [NONE] + cols
                idx = opts.index(guess) if guess in opts else 0
                sel = cc[n % 2].selectbox(label, opts, index=idx, key=f"map_{key}_{sigk}")
                mapping[key] = None if sel == NONE else sel
    n_map = sum(1 for k, v in mapping.items() if v)
    st.success(f"{n_map} dari {len(FIELDS)} field terpetakan.")

if not mapping.get("id"):
    for t in (tab_dash, tab_review, tab_out):
        with t:
            st.warning("Petakan kolom **ID RTUP / Assignment** pada tab *Data & Kolom* terlebih dahulu.")
    st.stop()

# ---------------------------------------------------------------- PERHITUNGAN
params_json = json.dumps(P, sort_keys=True)
map_t = tuple(sorted((k, v or "") for k, v in mapping.items()))
IDC = mapping["id"]
elig = eligible_mask(ss.raw, mapping, only_elig)
flags0, skipped = compute_flags(ss.raw[elig], map_t, params_json)      # baseline (data asli)
flags1, _ = compute_flags(ss.work[elig], map_t, params_json)           # kondisi data sekarang
set_now = set(zip(flags1["id"], flags1["kode"]))
orig_by_id = {k: g.reset_index(drop=True) for k, g in flags0.groupby("id", sort=False)}


def progress(i):
    g = orig_by_id.get(i)
    if g is None:
        return 0, 0
    done = sum(1 for k in g["kode"] if ss.just.get((i, k), {}).get("status"))
    return done, len(g)


def note_lines(i):
    lines = []
    g = orig_by_id.get(i)
    if g is not None:
        for k in g["kode"]:
            j = ss.just.get((i, k))
            if j:
                lines.append(f"#[{k}]: {j['status']}# {j['text']}".strip())
    for b in ss.bfl.get(i, []):
        lines.append(f"#{b['kode']}: {b['status']}# {b['text']}".strip())
    return lines


all_ids = list(dict.fromkeys(ss.raw.loc[elig, IDC].astype(str).str.strip()))
all_ids = [i for i in all_ids if i]
rows = []
for i, g in orig_by_id.items():
    d, t = progress(i)
    rows.append(dict(
        id=i, anomali=int((g["jenis"] == "Anomali").sum()), outlier=int((g["jenis"] == "Outlier").sum()),
        total=t, selesai=d, progres=d / t if t else 0.0,
        masih_terdeteksi=sum((i, k) in set_now for k in g["kode"]),
    ))
rtup_df = pd.DataFrame(rows, columns=["id", "anomali", "outlier", "total", "selesai", "progres", "masih_terdeteksi"])
if not rtup_df.empty:
    rtup_df = rtup_df.sort_values(["progres", "total"], ascending=[True, False]).reset_index(drop=True)

# ---------------------------------------------------------------- TAB 2
with tab_dash:
    n_el = int(elig.sum())
    if only_elig and mapping.get("status") and n_el < len(ss.raw):
        st.warning(f"{len(ss.raw) - n_el:,} baris dikecualikan karena status bukan APPROVED BY Pengawas / REVOKED BY Admin.")
    tot_flag = len(flags0)
    done_flag = sum(1 for (i, k) in zip(flags0["id"], flags0["kode"]) if ss.just.get((i, k), {}).get("status"))
    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Baris dianalisis", f"{n_el:,}")
    m2.metric("RTUP unik", f"{len(all_ids):,}")
    m3.metric("RTUP berflag", f"{len(rtup_df):,}")
    m4.metric("Total flag", f"{tot_flag:,}")
    m5.metric("Justifikasi selesai", f"{done_flag / tot_flag * 100:.0f}%" if tot_flag else "-")

    if skipped:
        with st.expander(f"⚠️ {len(skipped)} aturan dilewati"):
            st.dataframe(pd.DataFrame(skipped, columns=["Kode", "Alasan"]))

    if flags0.empty:
        st.success("Tidak ada anomali/outlier terdeteksi dengan pemetaan & parameter saat ini. 🎉")
    else:
        st.subheader("Ringkasan per kode flag")
        sm = flags0.groupby(["kode", "jenis", "deskripsi"]).agg(rtup=("id", "nunique")).reset_index()
        sm["masih_terdeteksi"] = sm["kode"].map(flags1.groupby("kode")["id"].nunique()).fillna(0).astype(int)
        sm["selesai_justifikasi"] = [
            sum(1 for i in flags0.loc[flags0["kode"] == k, "id"] if ss.just.get((i, k), {}).get("status"))
            for k in sm["kode"]
        ]
        st.bar_chart(sm.set_index("kode")["rtup"])
        st.dataframe(sm.rename(columns={"rtup": "jumlah_RTUP"}))

        st.subheader("Daftar RTUP prioritas cleaning")
        f1, f2 = st.columns([1, 2])
        jenis_f = f1.selectbox("Jenis", ["Semua", "Hanya anomali", "Hanya outlier"])
        cari = f2.text_input("Cari ID RTUP")
        v = rtup_df.copy()
        if jenis_f == "Hanya anomali":
            v = v[v["anomali"] > 0]
        elif jenis_f == "Hanya outlier":
            v = v[v["outlier"] > 0]
        if cari:
            v = v[v["id"].str.contains(cari, case=False, regex=False)]
        st.dataframe(
            v,
            column_config={"progres": st.column_config.ProgressColumn("progres", min_value=0, max_value=1, format="percent")},
        )

# ---------------------------------------------------------------- TAB 3
with tab_review:
    if ss.flash:
        st.success(ss.flash)
        ss.flash = None
    f1, f2 = st.columns([1, 2])
    mode = f1.radio("Tampilkan RTUP", ["Belum selesai", "Semua berflag", "Selesai"], horizontal=False)
    show_all = f1.checkbox("Sertakan RTUP tanpa flag", value=False)
    pool = []
    prog = {r.id: (r.selesai, r.total) for r in rtup_df.itertuples()}
    for i in (all_ids if show_all else list(rtup_df["id"])):
        d, t = prog.get(i, (0, 0))
        if mode == "Belum selesai" and t and d == t:
            continue
        if mode == "Selesai" and not (t and d == t):
            continue
        pool.append(i)
    if not pool:
        st.info("Tidak ada RTUP pada kategori ini.")
    else:
        def lbl(i):
            d, t = prog.get(i, (0, 0))
            return f"{'✅' if t and d == t else ('⏳' if t else '•')} {i} · {t} flag ({d} selesai)"

        sel = f2.selectbox("Pilih RTUP (ketik untuk mencari)", pool, format_func=lbl)
        g = orig_by_id.get(sel, pd.DataFrame(columns=flags0.columns))

        # ---- A. daftar flag
        st.subheader(f"RTUP {sel}")
        if g.empty:
            st.caption("RTUP ini tidak memiliki flag dari daftar anomali/outlier.")
        else:
            tbl = pd.DataFrame({
                "Kode": g["kode"], "Deskripsi": g["deskripsi"], "Info": g["info"],
                "Kondisi data sekarang": [
                    "⚠️ Masih terdeteksi" if (sel, k) in set_now else "✅ Tidak terdeteksi lagi" for k in g["kode"]
                ],
            })
            st.dataframe(tbl)

        # ---- B. edit data
        st.markdown("#### ① Perbaiki data (sesuai hasil konfirmasi)")
        wd = ss.work
        ridx = wd.index[wd[IDC].astype(str).str.strip() == sel]
        rel = ["r110"] + [f for k in g["kode"] for f in FLAG_FIELDS.get(k, [])]
        rel_cols = [mapping[f] for f in dict.fromkeys(rel) if mapping.get(f)]
        show_every = st.checkbox("Tampilkan semua kolom", value=False, key=f"allc_{sel}")
        vcols = list(wd.columns) if show_every else list(dict.fromkeys([IDC] + rel_cols))
        view = wd.loc[ridx, vcols].fillna("").astype(str)
        edited = st.data_editor(view, key=f"ed_{sel}_{ss.ver}", disabled=[IDC])
        if st.button("💾 Simpan perubahan data", key=f"save_{sel}_{ss.ver}"):
            nchg = 0
            for col in vcols:
                if col == IDC:
                    continue
                new = edited[col].fillna("").astype(str)
                diff = view[col] != new
                for r in diff[diff].index:
                    ss.log.append(dict(
                        waktu=datetime.now().strftime("%Y-%m-%d %H:%M:%S"), id=sel, baris=int(r),
                        kolom=col, nilai_lama=view.at[r, col], nilai_baru=new[r]))
                    ss.work.at[r, col] = new[r]
                    nchg += 1
            ss.ver += 1
            ss.flash = f"{nchg} sel berhasil diperbarui." if nchg else "Tidak ada perubahan."
            st.rerun()

        # ---- C. justifikasi
        if not g.empty:
            st.markdown("#### ② Justifikasi hasil konfirmasi")
            with st.form(f"form_{sel}"):
                for _, r in g.iterrows():
                    k = r["kode"]
                    saved = ss.just.get((sel, k), {})
                    st.markdown(f"**{k}** — {r['deskripsi']}" + (f"  \n_{r['info']}_" if r["info"] else ""))
                    a, b = st.columns([1, 2])
                    cur = saved.get("status", STATUS_PLACEHOLDER)
                    a.selectbox("Justifikasi", STATUS_OPTS, index=STATUS_OPTS.index(cur), key=f"st_{sel}_{k}")
                    b.text_area("Penjelasan hasil konfirmasi", value=saved.get("text", ""), key=f"tx_{sel}_{k}",
                                height=80, placeholder="Hasil konfirmasi kepada petugas atau responden, …")
                submitted = st.form_submit_button("💾 Simpan justifikasi")
            if submitted:
                for k in g["kode"]:
                    stt, txt = ss[f"st_{sel}_{k}"], ss[f"tx_{sel}_{k}"].strip()
                    if stt == STATUS_PLACEHOLDER:
                        ss.just.pop((sel, k), None)
                    else:
                        ss.just[(sel, k)] = {"status": stt, "text": txt}
                ss.flash = "Justifikasi tersimpan."
                st.rerun()
            for k in g["kode"]:
                j = ss.just.get((sel, k))
                if not j:
                    continue
                if j["status"] == "Diperbaiki" and (sel, k) in set_now:
                    st.warning(f"{k}: berstatus *Diperbaiki* tetapi data masih terdeteksi. Pastikan nilai sudah diubah.")
                if j["status"] != "Non Respon" and not j["text"]:
                    st.info(f"{k}: penjelasan hasil konfirmasi belum diisi.")

        # ---- D. flag B
        with st.expander("➕ Flagging di luar daftar (B_xx, oleh BPS Provinsi)"):
            nb = len(ss.bfl.get(sel, []))
            c1_, c2_ = st.columns([1, 2])
            num = c1_.text_input("Nomor anomali prov", key=f"bn_{sel}_{nb}", placeholder="01")
            bst = c2_.selectbox("Justifikasi", ["Diperbaiki", "Sesuai Kondisi Lapangan"], key=f"bs_{sel}_{nb}")
            btx = st.text_area("Penjelasan", key=f"bt_{sel}_{nb}", height=70)
            if st.button("Tambah flag B", key=f"badd_{sel}_{nb}"):
                digits_ = re.sub(r"\D", "", num)
                if not digits_:
                    st.error("Isi nomor anomali (angka).")
                else:
                    ss.bfl.setdefault(sel, []).append(dict(kode=f"B_{digits_.zfill(2)}", status=bst, text=btx.strip()))
                    st.rerun()
            for n_, b in enumerate(ss.bfl.get(sel, [])):
                x1, x2 = st.columns([6, 1])
                x1.code(f"#{b['kode']}: {b['status']}# {b['text']}".strip(), language=None)
                if x2.button("🗑️", key=f"bdel_{sel}_{n_}"):
                    ss.bfl[sel].pop(n_)
                    st.rerun()

        # ---- E. teks catatan
        st.markdown("#### ③ Teks untuk Blok Catatan Anomali (FASIH SM)")
        lines = note_lines(sel)
        if lines:
            st.code("\n".join(lines), language=None)
            st.caption("Salin (ikon di kanan atas) → buka assignment di FASIH SM → Review → Edit → tempel di Blok Catatan Anomali.")
        else:
            st.caption("Belum ada justifikasi yang disimpan untuk RTUP ini.")

# ---------------------------------------------------------------- TAB 4
with tab_out:
    ids_out = list(dict.fromkeys(list(orig_by_id.keys()) + list(ss.bfl.keys())))
    cat_rows = []
    for i in ids_out:
        d, t = progress(i)
        lines = note_lines(i)
        cat_rows.append(dict(
            id=i, jumlah_flag=t, justifikasi_selesai=d,
            status="Lengkap" if d == t else "Belum lengkap",
            catatan_anomali="\n".join(lines),
        ))
    cat_df = pd.DataFrame(cat_rows, columns=["id", "jumlah_flag", "justifikasi_selesai", "status", "catatan_anomali"])
    st.subheader("Catatan anomali per RTUP")
    if cat_df.empty:
        st.info("Belum ada data.")
    else:
        st.metric("RTUP lengkap", f"{(cat_df['status'] == 'Lengkap').sum()} / {len(cat_df)}")
        st.dataframe(cat_df)

    log_df = pd.DataFrame(ss.log, columns=["waktu", "id", "baris", "kolom", "nilai_lama", "nilai_baru"])
    flag_out = flags0.drop(columns=["rows"]).copy()
    if not flag_out.empty:
        flag_out["justifikasi"] = [ss.just.get((i, k), {}).get("status", "") for i, k in zip(flag_out["id"], flag_out["kode"])]
        flag_out["penjelasan"] = [ss.just.get((i, k), {}).get("text", "") for i, k in zip(flag_out["id"], flag_out["kode"])]
        flag_out["masih_terdeteksi"] = [(i, k) in set_now for i, k in zip(flag_out["id"], flag_out["kode"])]
    clean = ss.work.copy()
    cmap = dict(zip(cat_df["id"], cat_df["catatan_anomali"]))
    clean["catatan_anomali"] = clean[IDC].astype(str).str.strip().map(cmap).fillna("")

    st.subheader("Unduh")
    d1, d2, d3, d4 = st.columns(4)
    d1.download_button(
        "📥 Laporan lengkap (Excel)",
        to_excel({"Catatan_Anomali": cat_df, "Daftar_Flag": flag_out, "Log_Perubahan": log_df, "Data_Bersih": clean}),
        "cleaning_skp2026.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    d2.download_button("📥 Catatan anomali (CSV)", cat_df.to_csv(index=False).encode("utf-8-sig"),
                       "catatan_anomali_skp2026.csv", "text/csv")
    d3.download_button("📥 Data bersih (CSV)", clean.to_csv(index=False).encode("utf-8-sig"),
                       "data_bersih_skp2026.csv", "text/csv")
    d4.download_button("📥 Log perubahan (CSV)", log_df.to_csv(index=False).encode("utf-8-sig"),
                       "log_perubahan_skp2026.csv", "text/csv")
    if not log_df.empty:
        with st.expander(f"Log perubahan ({len(log_df)})"):
            st.dataframe(log_df)
    st.caption("Catatan: pengunggahan hasil perbaikan tetap dilakukan melalui web FASIH SM dengan akun admin.")

# ---------------------------------------------------------------- TAB 5
with tab_guide:
    st.markdown(
        """
### Tahapan cleaning data SKP 2026
1. Tim SKP BPS Kab/Kota (PJK Teknis dan/atau Admin) melakukan cleaning lewat **web FASIH SM dengan akun admin**.
   Data yang bisa diperbaiki: status **approved by pengawas** atau **revoked by admin BPS Kab/Kota**.
2. Prioritaskan RTUP (assignment) yang mengandung **anomali dan/atau outlier** (lihat dashboard SKP / FASIH-DASHBOARD).
3. Di FASIH SM: pilih RTUP → klik **Review** → klik **Edit** → isi hasil konfirmasi pada **Blok Catatan Anomali**.

### Jenis justifikasi
| Justifikasi | Format |
|---|---|
| Diperbaiki | `#[KODE FLAG]: Diperbaiki#` + penjelasan |
| Sudah sesuai kondisi lapangan | `#[KODE FLAG]: Sesuai Kondisi Lapangan#` + penjelasan |
| Responden tidak mau memberi konfirmasi | `#[KODE FLAG]: Non Respon#` |

Contoh: `#[A_08]: Sesuai Kondisi Lapangan# Hasil konfirmasi kepada petugas atau responden, upah pekerja diluar sektor pertanian benar dibawah 1000000 karena produksi sedang menurun.`

Bila ada beberapa flag pada satu RTUP, tulis satu baris catatan untuk masing-masing flag.

### Flagging di luar daftar
Untuk anomali/outlier yang dilakukan mandiri oleh BPS Provinsi: `#B_[Nomor anomali prov]: [Diperbaiki/sesuai kondisi lapangan]# [Penjelasan]`, mis. `#B_01: Diperbaiki# PENJELASAN`.
"""
    )
    st.subheader("Daftar anomali")
    st.dataframe(pd.DataFrame({"Kode": list(ANOMALI_DESC), "Deskripsi": list(ANOMALI_DESC.values())}))
    st.subheader("Daftar outlier")
    st.dataframe(pd.DataFrame({"Kode": list(OUTLIER_DESC), "Deskripsi": list(OUTLIER_DESC.values())}))
