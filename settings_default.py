# -*- coding: utf-8 -*-
"""
默认配置 - 当 settings.json 不存在时用这份初始化
所有可配置参数都在这里集中维护

安全提示: Excel 打开密码等敏感信息不出现在仓库中,
首次运行后请在主界面 [参数设置] 中填写, 或直接编辑 settings.json。
"""

import os

_ROOT = os.path.dirname(os.path.abspath(__file__))
_DEFAULT_TEMPLATE = os.path.join(
    _ROOT, "ROI combine Template",
    "Empty SFP254v1.5_Template_16-05-2024_single+multiprogram.xlsx")
_DEFAULT_OUTPUT = os.path.join(_ROOT, "output")

DEFAULT_SETTINGS = {
    "common": {
        # Excel 打开密码 - 请自行在 [参数设置] 中配置
        "open_password": ""
    },

    "tooling_modify": {
        "tooling_prefix": "Tooling",
        "tooling_flag_cell": "R121",
        "tooling_flag_val": "Cash Flow Projection for Tooling",
        "cbd_prefix": "CBD",
        "cbd_flag_cell": "C8",
        "cbd_flag_val": "SOP Year",
        "cbd_year_cell": "D8",
        "cbd_month_cell": "F8",
        "log_sheet_name": "_ChangeLog"
    },

    "tooling_rollback": {
        "log_sheet_name": "_ChangeLog"
    },

    "roi_combine": {
        "template_path": _DEFAULT_TEMPLATE,
        "output_folder": _DEFAULT_OUTPUT,
        "precer_prefix": "SFP254v1.5_CQSC-PreCER",
        "precer_flag_cell": "I7",
        "precer_flag_val": "Pre-CER / CER Financial Evaluation",
        "pi_prefix": "PI",
        "pi_flag_cell": "B3",
        "pi_flag_val": "Part Description",
        "pi_src_header_row": 3,
        "pi_src_data_row_start": 4,
        "pi_src_col_start": 2,
        "pi_src_col_end": 15,
        "pi_dst_header_row": 2,
        "pi_dst_data_row_start": 3,
        "pi_dst_col_start": 1,
        "pi_data_rows": 2,
        "pi_summary_sheet": "PI Summary",
        "cons_sheet": "SFP254v1.5_CONSOLIDATION",
        "cons_fill_cells": ["L11", "L12", "N365"],
        "cons_pic_range": "I352:Y392",
        "ppt_sheet": "PPT1",
        "ppt_anchor_cell": "A1",
        "bgn_sheet": "Bgn>>",
        "end_sheet": "<<End"
    },

    "cm_version": {
        "gp_prefix": "GP",
        "gp_flag_cell": "C4",
        "gp_flag_val": "Volume (pcs)",
        "cbd_prefix": "CBD",
        "cbd_flag_cell": "T320",
        "cbd_flag_val": "No.of DL Required per part",
        "cbd_flag_keywords": ["dl", "required", "part"],
        "cbd_m_range": "M321:M420",
        "cbd_m_factor": 0.7,
        "cbd_v_range": "V321:V420",
        "cbd_v_factor": 0.3,
        "gp_i26_value": 0.005,
        "gp_i26_format": "0.0%",
        "gp_i27_clear": True,
        "file_suffix": "+CM",
        "log_sheet_name": "_CM_ChangeLog"
    },

    "cm_rollback": {
        "log_sheet_name": "_CM_ChangeLog"
    }
}


# 用于参数设置 GUI 的字段元信息
SETTINGS_META = {
    "common": {
        "_label": "通用",
        "open_password": {"label": "Excel 打开密码", "type": "str"}
    },
    "tooling_modify": {
        "_label": "Tooling 修改",
        "tooling_prefix":   {"label": "Tooling Sheet 前缀", "type": "str"},
        "tooling_flag_cell":{"label": "Tooling 定位单元格", "type": "str"},
        "tooling_flag_val": {"label": "Tooling 定位内容",   "type": "str"},
        "cbd_prefix":       {"label": "CBD Sheet 前缀",     "type": "str"},
        "cbd_flag_cell":    {"label": "CBD 定位单元格",     "type": "str"},
        "cbd_flag_val":     {"label": "CBD 定位内容",       "type": "str"},
        "cbd_year_cell":    {"label": "CBD 年单元格",       "type": "str"},
        "cbd_month_cell":   {"label": "CBD 月单元格",       "type": "str"},
        "log_sheet_name":   {"label": "日志 Sheet 名",      "type": "str"}
    },
    "roi_combine": {
        "_label": "ROI 合并",
        "template_path":    {"label": "Template 路径",     "type": "path"},
        "output_folder":    {"label": "输出文件夹",         "type": "folder"},
        "precer_prefix":    {"label": "PreCER Sheet 前缀", "type": "str"},
        "precer_flag_cell": {"label": "PreCER 定位单元格", "type": "str"},
        "precer_flag_val":  {"label": "PreCER 定位内容",   "type": "str"},
        "pi_prefix":        {"label": "PI Sheet 前缀",     "type": "str"},
        "pi_flag_cell":     {"label": "PI 定位单元格",     "type": "str"},
        "pi_flag_val":      {"label": "PI 定位内容",       "type": "str"},
        "pi_summary_sheet": {"label": "PI Summary Sheet 名",   "type": "str"},
        "cons_sheet":       {"label": "CONSOLIDATION Sheet 名","type": "str"},
        "cons_pic_range":   {"label": "PPT 截图区域",       "type": "str"},
        "ppt_sheet":        {"label": "PPT 目标 Sheet",     "type": "str"},
        "ppt_anchor_cell":  {"label": "PPT 锚点单元格",     "type": "str"},
        "end_sheet":        {"label": "End Sheet 名",       "type": "str"}
    },
    "cm_version": {
        "_label": "To CM Version",
        "gp_prefix":         {"label": "GP Sheet 前缀", "type": "str"},
        "gp_flag_cell":      {"label": "GP 定位单元格", "type": "str"},
        "gp_flag_val":       {"label": "GP 定位内容", "type": "str"},
        "cbd_prefix":        {"label": "CBD Sheet 前缀", "type": "str"},
        "cbd_flag_cell":     {"label": "CBD 定位单元格", "type": "str"},
        "cbd_flag_val":      {"label": "CBD 定位内容", "type": "str"},
        "cbd_flag_keywords": {"label": "CBD 模糊关键词(英文逗号分隔)", "type": "list_str"},
        "cbd_m_range":       {"label": "M 列区域", "type": "str"},
        "cbd_m_factor":      {"label": "M 列系数", "type": "float"},
        "cbd_v_range":       {"label": "V 列区域", "type": "str"},
        "cbd_v_factor":      {"label": "V 列系数", "type": "float"},
        "gp_i26_value":      {"label": "I26 写入值", "type": "float"},
        "gp_i26_format":     {"label": "I26 数字格式", "type": "str"},
        "file_suffix":       {"label": "文件后缀", "type": "str"},
        "log_sheet_name":    {"label": "日志 Sheet 名", "type": "str"}
    }
}