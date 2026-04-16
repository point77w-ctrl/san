from __future__ import annotations

import json
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, simpledialog, ttk
from typing import Any

CONFIG_PATH = Path(__file__).with_name("tax_settings.json")

DEFAULT_CONFIG: dict[str, Any] = {
    "year": 2026,
    "premium_rates": {"overtime": 0.5, "night": 0.5, "holiday_le8": 0.5, "holiday_gt8": 1.0},
    "tax": {
        "local_income_tax_multiplier": 0.1,
        "income_tax_brackets": [
            {"up_to": 14000000, "rate": 0.06},
            {"up_to": 50000000, "rate": 0.15},
            {"up_to": 88000000, "rate": 0.24},
            {"up_to": 150000000, "rate": 0.35},
            {"up_to": 300000000, "rate": 0.38},
            {"up_to": 500000000, "rate": 0.40},
            {"up_to": 1000000000, "rate": 0.42},
            {"up_to": None, "rate": 0.45},
        ],
    },
    "insurance_rates": {
        "national_pension": 0.045,
        "health_insurance": 0.03545,
        "long_term_care": 0.00459,
        "employment_insurance": 0.009,
    },
    "personal_deduction_tabs": [
        {"id": "ded_1", "label": "개인공제1"},
        {"id": "ded_2", "label": "개인공제2"},
        {"id": "ded_3", "label": "개인공제3"},
    ],
    "column_label_overrides": {},
    "hidden_columns": [],
}

ID_COLUMNS = ("emp_no", "name", "rrn")
BASE_INPUT_COLUMNS = (
    "base_pay",
    "bonus",
    "hourly",
    "ot_h",
    "night_h",
    "holiday_le8_h",
    "holiday_gt8_h",
    "extra_allow",
)
RESULT_COLUMNS = ("gross", "income_tax", "local_tax", "ded_total", "net")
INSURANCE_COLUMNS = ("np", "hi", "ltc", "ei")


def load_config() -> dict[str, Any]:
    if not CONFIG_PATH.exists():
        CONFIG_PATH.write_text(json.dumps(DEFAULT_CONFIG, ensure_ascii=False, indent=2), encoding="utf-8")
        return json.loads(json.dumps(DEFAULT_CONFIG))
    with CONFIG_PATH.open("r", encoding="utf-8") as f:
        return json.load(f)


def save_config(config: dict[str, Any]) -> None:
    CONFIG_PATH.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")


def parse_int(value: str) -> int:
    raw = (value or "").replace(",", "").strip()
    return 0 if raw == "" else int(float(raw))


def parse_float(value: str) -> float:
    raw = (value or "").replace(",", "").strip()
    return 0.0 if raw == "" else float(raw)


def fmt_int(value: int) -> str:
    return f"{value:,}"


def fmt_negative(value: int) -> str:
    return f"-{abs(value):,}"


def krw(value: int) -> str:
    return f"{value:,.0f}원"


def progressive_tax(annual_income: int, brackets: list[dict[str, Any]]) -> int:
    tax = 0.0
    remain = annual_income
    lower = 0
    for b in brackets:
        upper = b.get("up_to")
        rate = float(b.get("rate", 0))
        if upper is None:
            taxable = max(remain, 0)
        else:
            span = max(int(upper) - lower, 0)
            taxable = min(max(remain, 0), span)
        tax += taxable * rate
        remain -= taxable
        if remain <= 0:
            break
        if upper is not None:
            lower = int(upper)
    return int(round(tax))


class PayrollMassCalculatorUI:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("개인별 급여 계산기 (자동 계산)")
        self.root.geometry("1880x980")
        self._configure_excel_like_style()

        self.config_data = load_config()
        self.personal_tabs: list[dict[str, str]] = self.config_data.get("personal_deduction_tabs", [])
        if not self.personal_tabs:
            self.personal_tabs = json.loads(json.dumps(DEFAULT_CONFIG["personal_deduction_tabs"]))
        self.column_label_overrides: dict[str, str] = {
            str(k): str(v) for k, v in self.config_data.get("column_label_overrides", {}).items()
        }
        self.hidden_columns: list[str] = [str(c) for c in self.config_data.get("hidden_columns", [])]

        self.drag_column_id: str | None = None
        self.last_clicked_header_col: str | None = None
        self.header_marker: tk.Canvas | None = None
        self.header_marker_under: tk.Canvas | None = None
        self.border_parts: list[tk.Canvas] = []
        self.active_cell: tuple[str, str] | None = None
        self.selection_anchor: tuple[str, str] | None = None
        self.selected_cells: set[tuple[str, str]] = set()
        self.multi_cell_parts: list[tk.Canvas] = []
        self.grid_line_parts: list[tk.Canvas] = []
        self.display_columns: list[str] = []

        self._build_fixed_section()
        self._build_bulk_fill_section()
        self._build_employee_section()
        self._build_summary_section()
        self._bind_global_context_fallback()

        self._load_config_to_screen()
        self._seed_sample_rows()
        self.calculate_all_rows()
        self.root.after_idle(self._autosize_all_columns)

    def _build_fixed_section(self) -> None:
        frame = ttk.LabelFrame(self.root, text="고정값(전체 공통)", padding=8)
        frame.pack(fill="x", padx=8, pady=6)

        self.vars: dict[str, tk.StringVar] = {
            "year": tk.StringVar(),
            "overtime_rate": tk.StringVar(),
            "night_rate": tk.StringVar(),
            "holiday_le8_rate": tk.StringVar(),
            "holiday_gt8_rate": tk.StringVar(),
            "local_income_multiplier": tk.StringVar(),
            "national_pension": tk.StringVar(),
            "health_insurance": tk.StringVar(),
            "long_term_care": tk.StringVar(),
            "employment_insurance": tk.StringVar(),
        }

        fields = [
            ("기준연도", "year"),
            ("연장가산율", "overtime_rate"),
            ("야간가산율", "night_rate"),
            ("휴일8h이내가산율", "holiday_le8_rate"),
            ("휴일8h초과가산율", "holiday_gt8_rate"),
            ("지방소득세배수(10%=0.1)", "local_income_multiplier"),
            ("국민연금율", "national_pension"),
            ("건강보험율", "health_insurance"),
            ("장기요양보험율", "long_term_care"),
            ("고용보험율", "employment_insurance"),
        ]

        for idx, (label, key) in enumerate(fields):
            row = idx // 6
            col = (idx % 6) * 2
            ttk.Label(frame, text=label).grid(row=row, column=col, padx=3, pady=2, sticky="w")
            e = ttk.Entry(frame, textvariable=self.vars[key], width=11)
            e.grid(row=row, column=col + 1, padx=3, pady=2)
            e.bind("<FocusOut>", lambda _evt: self.calculate_all_rows())

        ttk.Button(frame, text="설정 저장", command=self._save_fixed_to_config).grid(row=2, column=0, padx=6, pady=4, sticky="w")
        ttk.Button(frame, text="설정 불러오기", command=self._load_config_to_screen).grid(row=2, column=1, padx=6, pady=4, sticky="w")

    def _build_bulk_fill_section(self) -> None:
        frame = ttk.LabelFrame(self.root, text="일괄 입력/삭제", padding=8)
        frame.pack(fill="x", padx=8, pady=4)

        ttk.Label(frame, text="대상항목").pack(side="left", padx=4)
        self.bulk_col_var = tk.StringVar()
        self.bulk_col_combo = ttk.Combobox(frame, textvariable=self.bulk_col_var, width=20, state="readonly")
        self.bulk_col_combo.pack(side="left", padx=4)

        ttk.Label(frame, text="값").pack(side="left", padx=4)
        self.bulk_value_var = tk.StringVar(value="0")
        ttk.Entry(frame, textvariable=self.bulk_value_var, width=12).pack(side="left", padx=4)

        ttk.Button(frame, text="전체 행 동일값 입력", command=self._bulk_fill_column).pack(side="left", padx=6)
        ttk.Button(frame, text="해당 항목 전체 0", command=self._bulk_clear_column).pack(side="left", padx=6)

    def _build_employee_section(self) -> None:
        frame = ttk.LabelFrame(self.root, text="직원별 입력/결과", padding=8)
        frame.pack(fill="both", expand=True, padx=8, pady=6)

        self.emp_tree = ttk.Treeview(frame, show="headings", height=20, selectmode="extended", style="Excel.Treeview")
        self.emp_tree.pack(fill="both", expand=True)
        x_scroll = ttk.Scrollbar(frame, orient="horizontal", command=self.emp_tree.xview)
        x_scroll.pack(fill="x")
        self.emp_tree.configure(xscrollcommand=x_scroll.set)

        self.emp_tree.bind("<Double-1>", self._edit_cell)
        self.emp_tree.bind("<Double-Button-1>", self._on_header_double_click, add="+")
        self.emp_tree.bind("<Button-1>", self._single_click_select, add="+")
        self.emp_tree.bind("<KeyPress>", self._key_input_to_cell, add="+")
        self.emp_tree.bind("<Delete>", self._delete_selected_cells_to_zero, add="+")
        self._bind_context_menu(self.emp_tree)
        self.emp_tree.bind("<ButtonPress-1>", self._start_drag_tab, add="+")
        self.emp_tree.bind("<ButtonRelease-1>", self._drop_drag_tab, add="+")
        self.emp_tree.bind("<ButtonRelease-1>", self._track_header_click, add="+")
        self.emp_tree.bind("<ButtonRelease-1>", lambda _e: self._refresh_active_cell_overlay(), add="+")
        self.emp_tree.bind("<ButtonRelease-1>", lambda _e: self._refresh_multi_cell_overlays(), add="+")
        self.emp_tree.bind("<Configure>", lambda _e: self._redraw_grid_lines(), add="+")
        self.emp_tree.bind("<Configure>", lambda _e: self._refresh_active_cell_overlay(), add="+")
        self.emp_tree.bind("<Configure>", lambda _e: self._refresh_multi_cell_overlays(), add="+")
        self.emp_tree.bind("<ButtonRelease-1>", lambda _e: self._redraw_grid_lines(), add="+")

        button_row = ttk.Frame(frame)
        button_row.pack(fill="x", pady=4)
        ttk.Button(button_row, text="직원행 추가", command=self._add_employee_row).pack(side="left", padx=3)
        ttk.Button(button_row, text="선택행 삭제", command=self._delete_selected_rows).pack(side="left", padx=3)

        self._rebuild_employee_columns(preserve_rows=None)

    def _configure_excel_like_style(self) -> None:
        style = ttk.Style(self.root)
        style.configure("Excel.Treeview", rowheight=24, borderwidth=1, relief="solid")
        style.configure("Excel.Treeview.Heading", relief="solid")
        style.map("Excel.Treeview", background=[("selected", "#ffffff")], foreground=[("selected", "#000000")])

    def _bind_context_menu(self, widget: tk.Widget) -> None:
        widget.bind("<Button-3>", self._show_tab_context_menu_only)
        widget.bind("<ButtonPress-3>", self._show_tab_context_menu_only, add="+")
        widget.bind("<ButtonRelease-3>", self._show_tab_context_menu_only)
        widget.bind("<Button-2>", self._show_tab_context_menu_only, add="+")
        widget.bind("<ButtonPress-2>", self._show_tab_context_menu_only, add="+")
        widget.bind("<ButtonRelease-2>", self._show_tab_context_menu_only, add="+")

    def _bind_global_context_fallback(self) -> None:
        self.root.bind_all("<Button-3>", self._show_context_menu_global_fallback, add="+")
        self.root.bind_all("<ButtonRelease-3>", self._show_context_menu_global_fallback, add="+")
        self.root.bind_all("<Button-2>", self._show_context_menu_global_fallback, add="+")
        self.root.bind_all("<ButtonRelease-2>", self._show_context_menu_global_fallback, add="+")

    def _show_context_menu_global_fallback(self, _event: tk.Event[tk.Misc]) -> str | None:
        if not hasattr(self, "emp_tree"):
            return None
        px = self.emp_tree.winfo_pointerx()
        py = self.emp_tree.winfo_pointery()
        left = self.emp_tree.winfo_rootx()
        top = self.emp_tree.winfo_rooty()
        right = left + self.emp_tree.winfo_width()
        bottom = top + self.emp_tree.winfo_height()
        if not (left <= px <= right and top <= py <= bottom):
            return None
        x = int(px - left)
        y = int(py - top)
        self._show_header_context_menu_at(x, y)
        return "break"

    def _show_tab_context_menu_only(self, event: tk.Event[tk.Misc]) -> str | None:
        if isinstance(event.widget, ttk.Treeview):
            x = int(event.x)
            y = int(event.y)
        else:
            x = int(self.emp_tree.winfo_pointerx() - self.emp_tree.winfo_rootx())
            y = int(self.emp_tree.winfo_pointery() - self.emp_tree.winfo_rooty())
        self._show_header_context_menu_at(x, y)
        return "break"

    def _build_summary_section(self) -> None:
        frame = ttk.LabelFrame(self.root, text="요약", padding=8)
        frame.pack(fill="x", padx=8, pady=6)
        self.summary_var = tk.StringVar(value="입력하면 자동 계산됩니다.")
        ttk.Label(frame, textvariable=self.summary_var, font=("맑은 고딕", 11, "bold")).pack(anchor="w")

    def _single_click_select(self, event: tk.Event[tk.Misc]) -> None:
        if not isinstance(event.widget, ttk.Treeview):
            return
        tree = event.widget
        iid = tree.identify_row(event.y)
        col_id = tree.identify_column(event.x)
        self._clear_cell_border()
        if not iid:
            self.active_cell = None
            self.selected_cells.clear()
            self.selection_anchor = None
            self._clear_multi_cell_overlays()
            self._set_last_clicked_tab(None)
            return
        state = int(event.state)
        ctrl_or_shift = (state & 0x0001) or (state & 0x0004)
        if not ctrl_or_shift:
            tree.selection_set(iid)
        tree.focus(iid)
        col_name = self._column_name_from_col_id(col_id)
        self._set_last_clicked_tab(col_name)
        if col_id:
            clicked = (iid, col_id)
            if state & 0x0001 and self.selection_anchor:
                self.selected_cells = self._cells_in_rect(self.selection_anchor, clicked)
            elif state & 0x0004:
                if clicked in self.selected_cells:
                    self.selected_cells.remove(clicked)
                else:
                    self.selected_cells.add(clicked)
                self.selection_anchor = clicked
            else:
                self.selected_cells = {clicked}
                self.selection_anchor = clicked
            self.active_cell = clicked
            self._sync_row_selection_from_cells()
            self._refresh_active_cell_overlay()
            self._refresh_multi_cell_overlays()

    def _cells_in_rect(self, start: tuple[str, str], end: tuple[str, str]) -> set[tuple[str, str]]:
        row_ids = list(self.emp_tree.get_children())
        if not row_ids:
            return set()
        cols = [f"#{i+1}" for i in range(len(self.display_columns))]
        try:
            r1 = row_ids.index(start[0])
            r2 = row_ids.index(end[0])
            c1 = cols.index(start[1])
            c2 = cols.index(end[1])
        except ValueError:
            return {end}
        r_lo, r_hi = sorted((r1, r2))
        c_lo, c_hi = sorted((c1, c2))
        return {(row_ids[r], cols[c]) for r in range(r_lo, r_hi + 1) for c in range(c_lo, c_hi + 1)}

    def _sync_row_selection_from_cells(self) -> None:
        rows = sorted({iid for iid, _ in self.selected_cells}, key=lambda rid: list(self.emp_tree.get_children()).index(rid))
        if rows:
            self.emp_tree.selection_set(rows)

    def _clear_cell_border(self) -> None:
        for cv in self.border_parts:
            cv.place_forget()

    def _clear_multi_cell_overlays(self) -> None:
        for cv in self.multi_cell_parts:
            cv.destroy()
        self.multi_cell_parts.clear()

    def _draw_cell_outline(self, iid: str, col_id: str, color: str, thickness: int) -> None:
        bbox = self.emp_tree.bbox(iid, col_id)
        if not bbox:
            return
        x, y, w, h = bbox
        top = tk.Canvas(self.emp_tree, highlightthickness=0, bd=0, bg=color)
        bottom = tk.Canvas(self.emp_tree, highlightthickness=0, bd=0, bg=color)
        left = tk.Canvas(self.emp_tree, highlightthickness=0, bd=0, bg=color)
        right = tk.Canvas(self.emp_tree, highlightthickness=0, bd=0, bg=color)
        top.place(x=x, y=y, width=w, height=thickness)
        bottom.place(x=x, y=y + h - thickness, width=w, height=thickness)
        left.place(x=x, y=y, width=thickness, height=h)
        right.place(x=x + w - thickness, y=y, width=thickness, height=h)
        for cv in (top, bottom, left, right):
            self._bind_context_menu(cv)
            self._raise_widget(cv)
            self.multi_cell_parts.append(cv)

    def _refresh_multi_cell_overlays(self) -> None:
        self._clear_multi_cell_overlays()
        if not self.selected_cells:
            return
        row_ids = set(self.emp_tree.get_children())
        max_cells = 300
        for idx, (iid, col_id) in enumerate(self.selected_cells):
            if idx >= max_cells:
                break
            if iid not in row_ids:
                continue
            self._draw_cell_outline(iid, col_id, color="#7fb2ff", thickness=1)
        if self.active_cell and self.active_cell[0] in row_ids:
            self._draw_cell_outline(self.active_cell[0], self.active_cell[1], color="#2f7df6", thickness=2)

    def _refresh_active_cell_overlay(self) -> None:
        self._clear_cell_border()

    def _begin_edit_cell(self, iid: str, col_id: str, seed_text: str | None = None) -> None:
        col_name = self._column_name_from_col_id(col_id)
        if not col_name:
            return
        if col_name in INSURANCE_COLUMNS or col_name in RESULT_COLUMNS:
            return
        bbox = self.emp_tree.bbox(iid, col_id)
        if not bbox:
            return
        x, y, w, h = bbox
        old = self.emp_tree.set(iid, col_id)
        entry = ttk.Entry(self.emp_tree)
        entry.place(x=x, y=y, width=w, height=h)
        entry.insert(0, old if seed_text is None else seed_text)
        entry.focus_set()
        if seed_text is None:
            entry.selection_range(0, "end")
        else:
            entry.icursor("end")

        def save(_: tk.Event[tk.Misc] | None = None) -> None:
            raw = entry.get().strip()
            if col_name in ID_COLUMNS:
                self.emp_tree.set(iid, col_id, raw)
            elif self._is_input_numeric_column(col_name):
                self.emp_tree.set(iid, col_id, fmt_int(parse_int(raw)))
            else:
                self.emp_tree.set(iid, col_id, raw)
            entry.destroy()
            self.calculate_row(iid)
            self.update_summary()

        entry.bind("<Return>", save)
        entry.bind("<FocusOut>", save)

    def _key_input_to_cell(self, event: tk.Event[tk.Misc]) -> str | None:
        if not self.active_cell:
            return None
        iid, col_id = self.active_cell
        if event.keysym in {"Delete"}:
            return None
        ch = (event.char or "").strip()
        if not ch:
            return None
        if ch.isdigit() or ch in {"-", "."}:
            self._begin_edit_cell(iid, col_id, seed_text=ch)
            return "break"
        return None

    def _delete_selected_cells_to_zero(self, _event: tk.Event[tk.Misc]) -> str | None:
        targets = self.selected_cells or ({self.active_cell} if self.active_cell else set())
        if not targets:
            return None
        for iid, col_id in targets:
            col_name = self._column_name_from_col_id(col_id)
            if not col_name:
                continue
            if col_name in INSURANCE_COLUMNS or col_name in RESULT_COLUMNS:
                continue
            if col_name in ID_COLUMNS:
                self.emp_tree.set(iid, col_id, "")
            else:
                self.emp_tree.set(iid, col_id, "0")
            self.calculate_row(iid)
        self.update_summary()
        return "break"

    def _raise_widget(self, widget: tk.Widget) -> None:
        widget.tk.call("raise", widget._w)

    def _column_name_from_col_id(self, col_id: str) -> str | None:
        if not col_id:
            return None
        display_cols = list(self.emp_tree.cget("displaycolumns"))
        idx = int(col_id[1:]) - 1
        if idx < 0 or idx >= len(display_cols):
            return None
        return str(display_cols[idx])

    def _show_dotted_border(self, iid: str, col_id: str) -> None:
        if not iid or not col_id:
            return
        bbox = self.emp_tree.bbox(iid, col_id)
        if not bbox:
            return
        x, y, w, h = bbox

        if not self.border_parts:
            field_bg = ttk.Style(self.root).lookup("Treeview", "fieldbackground") or "#ffffff"
            for _ in range(4):
                c = tk.Canvas(self.emp_tree, highlightthickness=0, bd=0, bg=field_bg)
                self._bind_context_menu(c)
                self.border_parts.append(c)

        top, bottom, left, right = self.border_parts
        top.place(x=x, y=y, width=w, height=2)
        bottom.place(x=x, y=y + h - 2, width=w, height=2)
        left.place(x=x, y=y, width=2, height=h)
        right.place(x=x + w - 2, y=y, width=2, height=h)
        for cv in self.border_parts:
            self._raise_widget(cv)

        for cv in self.border_parts:
            cv.delete("all")
        top.create_line(0, 1, w, 1, fill="#2f7df6", width=2)
        bottom.create_line(0, 1, w, 1, fill="#2f7df6", width=2)
        left.create_line(1, 0, 1, h, fill="#2f7df6", width=2)
        right.create_line(1, 0, 1, h, fill="#2f7df6", width=2)

    def _current_column_defs(self) -> tuple[list[str], dict[str, str]]:
        personal_ids = [t["id"] for t in self.personal_tabs]
        columns = list(ID_COLUMNS) + list(BASE_INPUT_COLUMNS) + personal_ids + list(INSURANCE_COLUMNS) + list(RESULT_COLUMNS)

        headings = {
            "emp_no": "사번",
            "name": "이름",
            "rrn": "주민번호",
            "base_pay": "기본급",
            "bonus": "성과금",
            "hourly": "통상시급",
            "ot_h": "연장h",
            "night_h": "야간h",
            "holiday_le8_h": "휴일8↓h",
            "holiday_gt8_h": "휴일8↑h",
            "extra_allow": "기타수당",
            "np": "국민연금",
            "hi": "건강보험",
            "ltc": "장기요양",
            "ei": "고용보험",
            "gross": "총지급",
            "income_tax": "소득세",
            "local_tax": "지방세",
            "ded_total": "공제합계",
            "net": "실수령",
        }
        for tab in self.personal_tabs:
            headings[tab["id"]] = tab["label"]
        for col_name, label in self.column_label_overrides.items():
            if col_name in columns and col_name not in {t["id"] for t in self.personal_tabs}:
                headings[col_name] = label
        return columns, headings

    def _rebuild_employee_columns(self, preserve_rows: list[dict[str, str]] | None) -> None:
        columns, headings = self._current_column_defs()
        self.active_cell = None
        self.selected_cells.clear()
        self.selection_anchor = None
        self._clear_cell_border()
        self._clear_multi_cell_overlays()
        try:
            self.emp_tree.configure(displaycolumns="#all")
        except tk.TclError:
            pass
        self.emp_tree.configure(columns=columns)
        self.hidden_columns = [c for c in self.hidden_columns if c in columns]
        visible_columns = [c for c in columns if c not in self.hidden_columns]
        if not visible_columns:
            visible_columns = list(columns)
            self.hidden_columns.clear()
        if not self.display_columns:
            self.display_columns = list(visible_columns)
        else:
            existing = [c for c in self.display_columns if c in visible_columns]
            added = [c for c in visible_columns if c not in existing]
            self.display_columns = existing + added
        self.emp_tree.configure(displaycolumns=self.display_columns)
        self._show_header_indicator(self.last_clicked_header_col)

        for c in columns:
            anchor = "w" if c in ID_COLUMNS else "e"
            width = 95
            if c == "name":
                width = 110
            if c == "rrn":
                width = 130
            self.emp_tree.heading(c, text=headings.get(c, c))
            self.emp_tree.column(c, width=width, anchor=anchor)

        self._refresh_bulk_targets(headings)

        existing = preserve_rows or []
        for iid in self.emp_tree.get_children():
            self.emp_tree.delete(iid)
        self.active_cell = None
        self.selected_cells.clear()
        self.selection_anchor = None
        self._clear_cell_border()
        self._clear_multi_cell_overlays()

        for row in existing:
            values = [row.get(c, "0" if c not in ID_COLUMNS else "") for c in columns]
            self.emp_tree.insert("", "end", values=values)
        self._refresh_row_stripes()

    def _refresh_row_stripes(self) -> None:
        self.emp_tree.tag_configure("row_even", background="#ffffff")
        self.emp_tree.tag_configure("row_odd", background="#f7f9fc")
        for idx, iid in enumerate(self.emp_tree.get_children()):
            self.emp_tree.item(iid, tags=("row_even" if idx % 2 == 0 else "row_odd",))
        self._redraw_grid_lines()

    def _clear_grid_lines(self) -> None:
        for line in self.grid_line_parts:
            line.destroy()
        self.grid_line_parts.clear()

    def _redraw_grid_lines(self) -> None:
        if not hasattr(self, "emp_tree"):
            return
        self._clear_grid_lines()
        if not self.display_columns:
            return
        first_col = self.display_columns[0]
        total_width = sum(int(self.emp_tree.column(c, "width")) for c in self.display_columns)
        header_bottom = self._header_bottom_y()
        tree_h = max(header_bottom + 1, self.emp_tree.winfo_height())

        x = 0
        for col in self.display_columns[:-1]:
            x += int(self.emp_tree.column(col, "width"))
            vline = tk.Canvas(self.emp_tree, highlightthickness=0, bd=0, bg="#d6dbe1")
            vline.place(x=x, y=0, width=1, height=tree_h)
            self._bind_context_menu(vline)
            self.grid_line_parts.append(vline)

        for iid in self.emp_tree.get_children():
            bbox = self.emp_tree.bbox(iid, first_col)
            if not bbox:
                continue
            _, y, _, h = bbox
            hline = tk.Canvas(self.emp_tree, highlightthickness=0, bd=0, bg="#d6dbe1")
            hline.place(x=0, y=y + h - 1, width=max(1, total_width), height=1)
            self._bind_context_menu(hline)
            self.grid_line_parts.append(hline)

    def _refresh_bulk_targets(self, headings: dict[str, str]) -> None:
        target_cols = list(BASE_INPUT_COLUMNS) + [t["id"] for t in self.personal_tabs]
        display_items = [f"{headings[c]} ({c})" for c in target_cols]
        self.bulk_col_combo["values"] = display_items
        if display_items:
            self.bulk_col_combo.current(0)

    def _extract_col_id_from_display(self, display: str) -> str:
        if "(" in display and display.endswith(")"):
            return display.split("(")[-1].rstrip(")").strip()
        return display

    def _collect_rows_as_dict(self) -> list[dict[str, str]]:
        columns = self.emp_tree["columns"]
        rows: list[dict[str, str]] = []
        for iid in self.emp_tree.get_children():
            vals = self.emp_tree.item(iid, "values")
            row = {c: (vals[idx] if idx < len(vals) else "") for idx, c in enumerate(columns)}
            rows.append(row)
        return rows

    def _show_header_context_menu_at(self, x: int, y: int) -> None:
        menu = tk.Menu(self.root, tearoff=0)
        restore_candidates = [c for c in self.emp_tree["columns"] if c not in self.display_columns]
        if restore_candidates:
            restore_menu = tk.Menu(menu, tearoff=0)
            for col_name in restore_candidates:
                restore_label = self.emp_tree.heading(col_name, "text")
                restore_menu.add_command(label=f"{restore_label} ({col_name})", command=lambda c=col_name: self._restore_tab(c))
            menu.add_cascade(label="탭 복원", menu=restore_menu)

        col_id = self.emp_tree.identify_column(x)
        if not col_id:
            guessed = self._nearest_display_column(x)
            if guessed:
                idx = self.display_columns.index(guessed) + 1
                col_id = f"#{idx}"
        if col_id:
            col_name = self._column_name_from_col_id(col_id)
            if col_name:
                self._set_last_clicked_tab(col_name)
                menu.add_command(label="공제탭 추가(현재 탭 왼쪽)", command=lambda c=col_name: self._add_personal_tab(c))
                menu.add_command(label="탭 이름변경", command=lambda c=col_name: self._rename_tab(c))
                menu.add_command(label="탭 삭제", command=lambda c=col_name: self._delete_tab(c))
                menu.add_command(label="탭 왼쪽 이동", command=lambda c=col_name: self._move_tab(c, -1))
                menu.add_command(label="탭 오른쪽 이동", command=lambda c=col_name: self._move_tab(c, 1))
                menu.add_command(label="전체탭 너비 자동맞춤", command=self._autosize_all_columns)
        else:
            menu.add_command(label="공제탭 추가", command=lambda: self._add_personal_tab(self.last_clicked_header_col))

        try:
            menu.tk_popup(self.emp_tree.winfo_pointerx(), self.emp_tree.winfo_pointery())
        finally:
            menu.grab_release()

    def _nearest_display_column(self, x: int) -> str | None:
        columns = list(self.emp_tree.cget("displaycolumns"))
        pos = 0
        nearest_col = None
        nearest_dist = None
        for col in columns:
            width = int(self.emp_tree.column(col, "width"))
            center = pos + width / 2
            dist = abs(x - center)
            if nearest_dist is None or dist < nearest_dist:
                nearest_dist = dist
                nearest_col = col
            pos += width
        return nearest_col

    def _start_drag_tab(self, event: tk.Event[tk.Misc]) -> None:
        if self.emp_tree.identify_region(event.x, event.y) not in {"heading", "separator"}:
            self.drag_column_id = None
            return
        col_id = self.emp_tree.identify_column(event.x)
        if col_id:
            col_name = self._column_name_from_col_id(col_id)
        else:
            col_name = self._nearest_display_column(event.x)
        if col_name:
            self.drag_column_id = col_name
        else:
            self.drag_column_id = None

    def _drop_drag_tab(self, event: tk.Event[tk.Misc]) -> None:
        if not self.drag_column_id:
            return
        if self.emp_tree.identify_region(event.x, event.y) not in {"heading", "separator"}:
            self.drag_column_id = None
            return
        col_id = self.emp_tree.identify_column(event.x)
        if col_id:
            target_col = self._column_name_from_col_id(col_id)
        else:
            target_col = self._nearest_display_column(event.x)
        if not target_col:
            self.drag_column_id = None
            return
        self._reorder_display_column(self.drag_column_id, target_col)
        self.drag_column_id = None

    def _reorder_display_column(self, from_col: str, to_col: str) -> None:
        from_idx = next((i for i, c in enumerate(self.display_columns) if c == from_col), None)
        to_idx = next((i for i, c in enumerate(self.display_columns) if c == to_col), None)
        if from_idx is None or to_idx is None or from_idx == to_idx:
            return
        moving = self.display_columns.pop(from_idx)
        self.display_columns.insert(to_idx, moving)
        self.emp_tree.configure(displaycolumns=self.display_columns)

    def _move_tab(self, col_id: str, direction: int) -> None:
        idx = next((i for i, c in enumerate(self.display_columns) if c == col_id), None)
        if idx is None:
            return
        new_idx = idx + direction
        if new_idx < 0 or new_idx >= len(self.display_columns):
            return
        item = self.display_columns.pop(idx)
        self.display_columns.insert(new_idx, item)
        self.emp_tree.configure(displaycolumns=self.display_columns)

    def _add_personal_tab(self, anchor_col: str | None = None) -> None:
        rows = self._collect_rows_as_dict()
        next_idx = 1
        existing_ids = {t["id"] for t in self.personal_tabs}
        while f"ded_{next_idx}" in existing_ids:
            next_idx += 1
        new_id = f"ded_{next_idx}"
        new_tab = {"id": new_id, "label": f"개인공제{next_idx}"}
        anchor_idx = next((i for i, tab in enumerate(self.personal_tabs) if tab["id"] == anchor_col), None)
        if anchor_idx is None:
            self.personal_tabs.append(new_tab)
        else:
            self.personal_tabs.insert(anchor_idx, new_tab)
        self._rebuild_employee_columns(rows)
        if anchor_col and anchor_col in self.display_columns and new_id in self.display_columns:
            self.display_columns.remove(new_id)
            left_idx = self.display_columns.index(anchor_col)
            self.display_columns.insert(left_idx, new_id)
            self.emp_tree.configure(displaycolumns=self.display_columns)
        self._set_last_clicked_tab(new_id)
        self.calculate_all_rows()

    def _set_last_clicked_tab(self, col_name: str | None) -> None:
        self.last_clicked_header_col = col_name
        self._show_header_indicator(col_name)

    def _header_bottom_y(self) -> int:
        first = next(iter(self.emp_tree.get_children()), None)
        if not first:
            return 24
        first_col = next(iter(self.display_columns), None)
        if not first_col:
            return 24
        bbox = self.emp_tree.bbox(first, first_col)
        if not bbox:
            return 24
        return max(20, int(bbox[1]))

    def _show_header_indicator(self, col_name: str | None) -> None:
        if self.header_marker is None:
            self.header_marker = tk.Canvas(self.emp_tree, highlightthickness=0, bd=0)
            self._bind_context_menu(self.header_marker)
        if self.header_marker_under is None:
            self.header_marker_under = tk.Canvas(self.emp_tree, highlightthickness=0, bd=0)
            self._bind_context_menu(self.header_marker_under)
        if not col_name or col_name not in self.display_columns:
            self.header_marker.place_forget()
            self.header_marker_under.place_forget()
            return
        x = 0
        for c in self.display_columns:
            w = int(self.emp_tree.column(c, "width"))
            if c == col_name:
                y = self._header_bottom_y() - 3
                self.header_marker.place(x=x, y=max(0, y), width=max(8, w), height=3)
                self.header_marker.delete("all")
                self.header_marker.create_line(0, 1, w, 1, fill="#1f62ff", width=3)
                self._raise_widget(self.header_marker)
                under_y = self._header_bottom_y()
                self.header_marker_under.place(x=x, y=max(0, under_y), width=max(8, w), height=3)
                self.header_marker_under.delete("all")
                self.header_marker_under.create_line(0, 1, w, 1, fill="#1f62ff", width=3)
                self._raise_widget(self.header_marker_under)
                return
            x += w
        self.header_marker.place_forget()
        self.header_marker_under.place_forget()

    def _track_header_click(self, event: tk.Event[tk.Misc]) -> None:
        if self.emp_tree.identify_region(event.x, event.y) not in {"heading", "separator"}:
            return
        col_id = self.emp_tree.identify_column(event.x)
        col_name = self._column_name_from_col_id(col_id) if col_id else self._nearest_display_column(event.x)
        if col_name:
            self._set_last_clicked_tab(col_name)

    def _rename_tab(self, col_id: str) -> None:
        current_label = str(self.emp_tree.heading(col_id, "text") or col_id)
        new_label = simpledialog.askstring("탭 이름변경", "새 탭 이름", initialvalue=current_label, parent=self.root)
        if not new_label:
            return
        cleaned = new_label.strip()
        if not cleaned:
            return
        rows = self._collect_rows_as_dict()
        target = next((t for t in self.personal_tabs if t["id"] == col_id), None)
        if target:
            target["label"] = cleaned
        else:
            self.column_label_overrides[col_id] = cleaned
        self._rebuild_employee_columns(rows)

    def _delete_tab(self, col_id: str) -> None:
        if col_id not in self.emp_tree["columns"]:
            return
        rows = self._collect_rows_as_dict()
        target = next((t for t in self.personal_tabs if t["id"] == col_id), None)
        if target:
            if len(self.personal_tabs) <= 1:
                messagebox.showwarning("삭제 불가", "최소 1개의 개인공제 탭은 남겨야 합니다.")
                return
            self.personal_tabs = [t for t in self.personal_tabs if t["id"] != col_id]
        else:
            if col_id in ID_COLUMNS:
                messagebox.showwarning("삭제 불가", "사번/이름/주민번호 탭은 삭제할 수 없습니다.")
                return
            if len(self.display_columns) <= 1:
                messagebox.showwarning("삭제 불가", "최소 1개의 탭은 남겨야 합니다.")
                return
            if col_id not in self.hidden_columns:
                self.hidden_columns.append(col_id)
        self._rebuild_employee_columns(rows)
        if self.last_clicked_header_col == col_id:
            self._set_last_clicked_tab(None)
        self.calculate_all_rows()

    def _restore_tab(self, col_id: str) -> None:
        if col_id in self.hidden_columns:
            self.hidden_columns.remove(col_id)
        if col_id not in self.display_columns:
            self.display_columns.append(col_id)
        self.emp_tree.configure(displaycolumns=self.display_columns)

    def _load_config_to_screen(self) -> None:
        self.config_data = load_config()
        self.personal_tabs = self.config_data.get("personal_deduction_tabs", json.loads(json.dumps(DEFAULT_CONFIG["personal_deduction_tabs"])))
        for tab in self.personal_tabs:
            tab["label"] = str(tab.get("label", "")).lstrip("-").strip() or "개인공제"
        self.column_label_overrides = {
            str(k): str(v) for k, v in self.config_data.get("column_label_overrides", {}).items()
        }
        self.hidden_columns = [str(c) for c in self.config_data.get("hidden_columns", [])]

        premium = self.config_data.get("premium_rates", {})
        tax = self.config_data.get("tax", {})
        ins = self.config_data.get("insurance_rates", {})

        self.vars["year"].set(str(self.config_data.get("year", 2026)))
        self.vars["overtime_rate"].set(str(premium.get("overtime", 0.5)))
        self.vars["night_rate"].set(str(premium.get("night", 0.5)))
        self.vars["holiday_le8_rate"].set(str(premium.get("holiday_le8", 0.5)))
        self.vars["holiday_gt8_rate"].set(str(premium.get("holiday_gt8", 1.0)))
        self.vars["local_income_multiplier"].set(str(tax.get("local_income_tax_multiplier", 0.1)))

        self.vars["national_pension"].set(str(ins.get("national_pension", 0.045)))
        self.vars["health_insurance"].set(str(ins.get("health_insurance", 0.03545)))
        self.vars["long_term_care"].set(str(ins.get("long_term_care", 0.00459)))
        self.vars["employment_insurance"].set(str(ins.get("employment_insurance", 0.009)))

        rows = self._collect_rows_as_dict()
        self._rebuild_employee_columns(rows)
        self.calculate_all_rows()

    def _save_fixed_to_config(self) -> None:
        try:
            self.config_data["year"] = parse_int(self.vars["year"].get())
            self.config_data["premium_rates"] = {
                "overtime": parse_float(self.vars["overtime_rate"].get()),
                "night": parse_float(self.vars["night_rate"].get()),
                "holiday_le8": parse_float(self.vars["holiday_le8_rate"].get()),
                "holiday_gt8": parse_float(self.vars["holiday_gt8_rate"].get()),
            }
            self.config_data.setdefault("tax", {})["local_income_tax_multiplier"] = parse_float(self.vars["local_income_multiplier"].get())
            self.config_data["insurance_rates"] = {
                "national_pension": parse_float(self.vars["national_pension"].get()),
                "health_insurance": parse_float(self.vars["health_insurance"].get()),
                "long_term_care": parse_float(self.vars["long_term_care"].get()),
                "employment_insurance": parse_float(self.vars["employment_insurance"].get()),
            }
            self.config_data["personal_deduction_tabs"] = self.personal_tabs
            self.config_data["column_label_overrides"] = self.column_label_overrides
            self.config_data["hidden_columns"] = self.hidden_columns
            save_config(self.config_data)
            messagebox.showinfo("저장", "설정을 저장했습니다.")
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("저장 오류", str(e))

    def _seed_sample_rows(self) -> None:
        if self.emp_tree.get_children():
            return
        columns, _ = self._current_column_defs()
        for i in range(1, 6):
            row: dict[str, str] = {c: "0" for c in columns if c not in ID_COLUMNS}
            row["emp_no"] = f"E{i:04d}"
            row["name"] = f"직원{i}"
            row["rrn"] = "900101-1234567"
            row["base_pay"] = "2,500,000"
            row["hourly"] = "12,000"
            values = [row.get(c, "") for c in columns]
            self.emp_tree.insert("", "end", values=values)
        self._refresh_row_stripes()

    def _add_employee_row(self) -> None:
        columns, _ = self._current_column_defs()
        row = {c: "0" for c in columns if c not in ID_COLUMNS}
        row.update({"emp_no": "", "name": "", "rrn": ""})
        self.emp_tree.insert("", "end", values=[row.get(c, "") for c in columns])
        self._refresh_row_stripes()

    def _delete_selected_rows(self) -> None:
        for iid in self.emp_tree.selection():
            self.emp_tree.delete(iid)
        self._refresh_row_stripes()
        self.calculate_all_rows()

    def _bulk_fill_column(self) -> None:
        try:
            selected = self.bulk_col_var.get().strip()
            col_id = self._extract_col_id_from_display(selected)
            value = fmt_int(parse_int(self.bulk_value_var.get()))
            for iid in self.emp_tree.get_children():
                self.emp_tree.set(iid, col_id, value)
            self.calculate_all_rows()
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("일괄입력 오류", str(e))

    def _bulk_clear_column(self) -> None:
        selected = self.bulk_col_var.get().strip()
        col_id = self._extract_col_id_from_display(selected)
        for iid in self.emp_tree.get_children():
            self.emp_tree.set(iid, col_id, "0")
        self.calculate_all_rows()

    def _is_input_numeric_column(self, col: str) -> bool:
        return col in BASE_INPUT_COLUMNS or any(col == t["id"] for t in self.personal_tabs)

    def _edit_cell(self, event: tk.Event[tk.Misc]) -> None:
        tree = event.widget
        if not isinstance(tree, ttk.Treeview):
            return

        iid = tree.identify_row(event.y)
        col_id = tree.identify_column(event.x)
        region = tree.identify_region(event.x, event.y)
        if region in {"heading", "separator"}:
            return
        if not iid or not col_id:
            return

        col_name = self._column_name_from_col_id(col_id)
        if not col_name:
            return
        if col_name in INSURANCE_COLUMNS or col_name in RESULT_COLUMNS:
            return

        self.active_cell = (iid, col_id)
        self.selected_cells = {(iid, col_id)}
        self.selection_anchor = (iid, col_id)
        self._begin_edit_cell(iid, col_id)

    def _on_header_double_click(self, event: tk.Event[tk.Misc]) -> str | None:
        region = self.emp_tree.identify_region(event.x, event.y)
        if region not in {"heading", "separator"}:
            return None
        self._autosize_all_columns()
        return "break"

    def _autosize_column(self, col_name: str) -> None:
        header = str(self.emp_tree.heading(col_name, "text"))
        header_len = len(header)
        value_len = 0
        for iid in self.emp_tree.get_children():
            value = str(self.emp_tree.set(iid, col_name))
            if len(value) > value_len:
                value_len = len(value)
        use_len = max(header_len, value_len)
        width = max(80, min(220, use_len * 9 + 8))
        self.emp_tree.column(col_name, width=width)

    def _autosize_all_columns(self) -> None:
        for col_name in self.display_columns:
            self._autosize_column(col_name)

    def calculate_row(self, iid: str) -> None:
        premium = {
            "overtime": parse_float(self.vars["overtime_rate"].get()),
            "night": parse_float(self.vars["night_rate"].get()),
            "holiday_le8": parse_float(self.vars["holiday_le8_rate"].get()),
            "holiday_gt8": parse_float(self.vars["holiday_gt8_rate"].get()),
        }
        local_multiplier = parse_float(self.vars["local_income_multiplier"].get())
        brackets = self.config_data.get("tax", {}).get("income_tax_brackets", DEFAULT_CONFIG["tax"]["income_tax_brackets"])

        ins_np_rate = parse_float(self.vars["national_pension"].get())
        ins_hi_rate = parse_float(self.vars["health_insurance"].get())
        ins_ltc_rate = parse_float(self.vars["long_term_care"].get())
        ins_ei_rate = parse_float(self.vars["employment_insurance"].get())

        base_pay = parse_int(self.emp_tree.set(iid, "base_pay"))
        bonus = parse_int(self.emp_tree.set(iid, "bonus"))
        hourly = parse_int(self.emp_tree.set(iid, "hourly"))
        ot_h = parse_float(self.emp_tree.set(iid, "ot_h"))
        night_h = parse_float(self.emp_tree.set(iid, "night_h"))
        holiday_le8_h = parse_float(self.emp_tree.set(iid, "holiday_le8_h"))
        holiday_gt8_h = parse_float(self.emp_tree.set(iid, "holiday_gt8_h"))
        extra_allow = parse_int(self.emp_tree.set(iid, "extra_allow"))

        personal_ded_total = 0
        for tab in self.personal_tabs:
            personal_ded_total += abs(parse_int(self.emp_tree.set(iid, tab["id"])))

        overtime_allow = int(round(hourly * ot_h * premium["overtime"]))
        night_allow = int(round(hourly * night_h * premium["night"]))
        holiday_allow = int(round(hourly * holiday_le8_h * premium["holiday_le8"] + hourly * holiday_gt8_h * premium["holiday_gt8"]))

        gross = base_pay + bonus + extra_allow + overtime_allow + night_allow + holiday_allow
        income_tax = int(round(progressive_tax(gross * 12, brackets) / 12))
        local_tax = int(round(income_tax * local_multiplier))

        np = int(round(gross * ins_np_rate))
        hi = int(round(gross * ins_hi_rate))
        ltc = int(round(gross * ins_ltc_rate))
        ei = int(round(gross * ins_ei_rate))

        ded_total = personal_ded_total + np + hi + ltc + ei + income_tax + local_tax
        net = gross - ded_total

        self.emp_tree.set(iid, "np", fmt_negative(np))
        self.emp_tree.set(iid, "hi", fmt_negative(hi))
        self.emp_tree.set(iid, "ltc", fmt_negative(ltc))
        self.emp_tree.set(iid, "ei", fmt_negative(ei))

        self.emp_tree.set(iid, "gross", fmt_int(gross))
        self.emp_tree.set(iid, "income_tax", fmt_int(income_tax))
        self.emp_tree.set(iid, "local_tax", fmt_int(local_tax))
        self.emp_tree.set(iid, "ded_total", fmt_int(ded_total))
        self.emp_tree.set(iid, "net", fmt_int(net))

    def calculate_all_rows(self) -> None:
        for iid in self.emp_tree.get_children():
            self.calculate_row(iid)
        self.update_summary()

    def update_summary(self) -> None:
        total_gross = 0
        total_net = 0
        count = 0
        for iid in self.emp_tree.get_children():
            total_gross += parse_int(self.emp_tree.set(iid, "gross"))
            total_net += parse_int(self.emp_tree.set(iid, "net"))
            count += 1
        self.summary_var.set(f"직원 {count}명 | 총지급합계 {krw(total_gross)} | 실수령합계 {krw(total_net)}")


def main() -> None:
    root = tk.Tk()
    PayrollMassCalculatorUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
