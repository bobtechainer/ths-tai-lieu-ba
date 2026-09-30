# -*- coding: utf-8 -*-
"""Gom kết quả trace: mỗi đầu mục 39-75 -> danh sách nội dung + trạng thái."""
import io, json
from collections import OrderedDict

import build as B
import map2 as M
from build import CO, KHAC, DANG, KHONG

NHOM = {"truong": range(39, 44), "lop": range(44, 49), "nhansu": range(49, 54),
        "hocsinh": range(54, 59), "pccn": range(59, 64), "ttmh": range(64, 68),
        "lopmon": range(68, 72), "pccm": range(72, 76)}
CUA = {k: g for g, rg in NHOM.items() for k in rg}

GHI_DT = "Hệ thống không có phần khai báo điểm trường"
GHI_VKK = "Chỉ có ô tích Vùng đặc biệt khó khăn, không khai được danh mục vùng khó khăn"
GHI_HT = "Hệ thống không có phần khai báo hạ tầng Internet, máy tính và thiết bị"
NGOAI_LE_HT = {"KINH_TUYEN": (CO, "Trường Học > Kinh tuyến"),
               "VI_TUYEN": (CO, "Trường Học > Vĩ tuyến")}


def _khoi(path):
    if "DIEM_TRUONG" in path:
        return "DT"
    if "VUNG_KHO_KHAN" in path:
        return "VKK"
    if "HA_TANG_INTERNET" in path:
        return "HT"
    if "THONG_TIN_CAP_NHAT" in path:
        return "TT"
    return ""


def _tra(stt, path, name):
    g = CUA[stt]
    if g == "truong":
        kh = _khoi(path)
        ng = B.GIAI_TRUONG.get(name, name)
        if kh == "DT":
            return KHONG, GHI_DT, B.NHOM_DT, ng
        if kh == "VKK":
            return KHONG, GHI_VKK, B.NHOM_DT, ng
        if kh == "HT":
            if name in NGOAI_LE_HT:
                st, wh = NGOAI_LE_HT[name]
                return st, wh, B.NHOM_HT, ng
            return KHONG, GHI_HT, B.NHOM_HT, ng
        v = B.M_TRUONG[(kh, name)]
        return v[0], v[1], v[2], ng
    if g == "lop":
        v = B.M_LOP[name]
        return v[0], v[1], v[2], B.GIAI_LOP.get(name, name)
    if g == "nhansu":
        v = M.M_NS[name]
        return v[0], v[1], v[2], M.GIAI_NS.get(name, name)
    if g == "hocsinh":
        v = M.M_HS[name]
        return v[0], v[1], v[2], M.GIAI_HS.get(name, name)
    v = M.M_NHO[name]
    return v[0], v[1], v[2], M.GIAI_NHO.get(name, name)


def build_rows(path="_rs/api_fields.json"):
    api = json.load(io.open(path, encoding="utf-8"))
    rows = OrderedDict()
    for k in range(39, 76):
        it = api[str(k)]
        fs = []
        for f in it["fields"]:
            st, wh, nh, ng = _tra(k, f["path"], f["name"])
            fs.append({"name": f["name"], "bb": f["bb"], "nghia": ng,
                       "st": st, "wh": wh, "nhom": nh})
        rows[k] = {"ten": it["ten"], "fs": fs}
    return rows
