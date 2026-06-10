"""Shared shadcn-inspired desktop design tokens and ttk styles."""

from __future__ import annotations


COLORS = {
    "background": "#F8FAFC",
    "surface": "#FFFFFF",
    "surface_alt": "#F1F5F9",
    "surface_hover": "#F8FAFC",
    "primary": "#1E3A5F",
    "primary_hover": "#172F4D",
    "primary_dark": "#0F172A",
    "primary_soft": "#E8F0F8",
    "text": "#0F172A",
    "text_secondary": "#475569",
    "text_muted": "#94A3B8",
    "border": "#E2E8F0",
    "border_strong": "#CBD5E1",
    "divider": "#E2E8F0",
    "success": "#15803D",
    "success_soft": "#F0FDF4",
    "warning": "#A16207",
    "warning_soft": "#FFFBEB",
    "danger": "#B91C1C",
    "danger_soft": "#FEF2F2",
}

FONT_FAMILY = "Microsoft YaHei UI"


def configure_office_theme(style):
    """Register the compact component set shared by all desktop windows."""
    style.configure(".", font=(FONT_FAMILY, 10))
    style.configure("Office.TFrame", background=COLORS["background"])
    style.configure("Content.TFrame", background=COLORS["background"])
    style.configure("Surface.TFrame", background=COLORS["surface"])
    style.configure("Card.TFrame", background=COLORS["surface"])
    style.configure("Toolbar.TFrame", background=COLORS["surface"])
    style.configure("Sidebar.TFrame", background=COLORS["surface"])
    style.configure(
        "SidebarTitle.TLabel",
        background=COLORS["surface"],
        foreground=COLORS["text"],
        font=(FONT_FAMILY, 13, "bold"),
    )
    style.configure(
        "SidebarText.TLabel",
        background=COLORS["surface"],
        foreground=COLORS["text_muted"],
        font=(FONT_FAMILY, 9),
    )
    style.configure(
        "Eyebrow.TLabel",
        background=COLORS["background"],
        foreground=COLORS["text_muted"],
        font=(FONT_FAMILY, 9, "bold"),
    )
    style.configure(
        "PageTitle.TLabel",
        background=COLORS["background"],
        foreground=COLORS["text"],
        font=(FONT_FAMILY, 20, "bold"),
    )
    style.configure(
        "DialogTitle.TLabel",
        background=COLORS["background"],
        foreground=COLORS["text"],
        font=(FONT_FAMILY, 16, "bold"),
    )
    style.configure(
        "PageHint.TLabel",
        background=COLORS["background"],
        foreground=COLORS["text_secondary"],
        font=(FONT_FAMILY, 10),
    )
    style.configure(
        "CardTitle.TLabel",
        background=COLORS["surface"],
        foreground=COLORS["text"],
        font=(FONT_FAMILY, 11, "bold"),
    )
    style.configure(
        "Muted.TLabel",
        background=COLORS["surface"],
        foreground=COLORS["text_muted"],
        font=(FONT_FAMILY, 9),
    )
    style.configure(
        "Badge.TLabel",
        background=COLORS["surface_alt"],
        foreground=COLORS["text_secondary"],
        padding=(8, 3),
        font=(FONT_FAMILY, 9, "bold"),
    )
    style.configure(
        "Nav.TButton",
        background=COLORS["surface"],
        foreground=COLORS["text_secondary"],
        borderwidth=0,
        anchor="w",
        padding=(14, 10),
        font=(FONT_FAMILY, 10),
    )
    style.map(
        "Nav.TButton",
        background=[("active", COLORS["surface_alt"]), ("pressed", COLORS["surface_alt"])],
        foreground=[("active", COLORS["text"])],
    )
    style.configure(
        "NavActive.TButton",
        background=COLORS["primary_soft"],
        foreground=COLORS["primary"],
        borderwidth=0,
        anchor="w",
        padding=(14, 10),
        font=(FONT_FAMILY, 10, "bold"),
    )
    style.map(
        "NavActive.TButton",
        background=[("active", "#DCE8F4"), ("pressed", "#DCE8F4")],
    )
    style.configure("Workspace.TNotebook", background=COLORS["background"], borderwidth=0)
    style.layout("Workspace.TNotebook.Tab", [])
    style.configure(
        "Primary.TButton",
        background=COLORS["primary"],
        foreground="#FFFFFF",
        bordercolor=COLORS["primary"],
        lightcolor=COLORS["primary"],
        darkcolor=COLORS["primary"],
        padding=(14, 8),
        font=(FONT_FAMILY, 10, "bold"),
    )
    style.map(
        "Primary.TButton",
        background=[("active", COLORS["primary_hover"]), ("pressed", COLORS["primary_dark"])],
        bordercolor=[("active", COLORS["primary_hover"])],
    )
    style.configure(
        "Secondary.TButton",
        background=COLORS["surface"],
        foreground=COLORS["text"],
        bordercolor=COLORS["border_strong"],
        lightcolor=COLORS["border_strong"],
        darkcolor=COLORS["border_strong"],
        padding=(12, 7),
    )
    style.map(
        "Secondary.TButton",
        background=[("active", COLORS["surface_alt"]), ("pressed", COLORS["surface_alt"])],
    )
    style.configure(
        "Ghost.TButton",
        background=COLORS["background"],
        foreground=COLORS["text_secondary"],
        borderwidth=0,
        padding=(10, 7),
    )
    style.map(
        "Ghost.TButton",
        background=[("active", COLORS["surface_alt"])],
        foreground=[("active", COLORS["text"])],
    )
    style.configure(
        "Danger.TButton",
        background=COLORS["surface"],
        foreground=COLORS["danger"],
        bordercolor=COLORS["border_strong"],
        padding=(12, 7),
    )
    style.map("Danger.TButton", background=[("active", COLORS["danger_soft"])])
    style.configure(
        "Card.TLabelframe",
        background=COLORS["surface"],
        bordercolor=COLORS["border"],
        lightcolor=COLORS["border"],
        darkcolor=COLORS["border"],
        borderwidth=1,
        relief="solid",
    )
    style.configure(
        "Card.TLabelframe.Label",
        background=COLORS["surface"],
        foreground=COLORS["text"],
        font=(FONT_FAMILY, 10, "bold"),
        padding=(2, 0),
    )
    style.configure(
        "Treeview",
        background=COLORS["surface"],
        fieldbackground=COLORS["surface"],
        foreground=COLORS["text"],
        rowheight=32,
        bordercolor=COLORS["border"],
        lightcolor=COLORS["border"],
        darkcolor=COLORS["border"],
        borderwidth=1,
    )
    style.map(
        "Treeview",
        background=[("selected", COLORS["primary_soft"])],
        foreground=[("selected", COLORS["primary"])],
    )
    style.configure(
        "Treeview.Heading",
        background=COLORS["surface_alt"],
        foreground=COLORS["text_secondary"],
        padding=(9, 8),
        relief="flat",
        font=(FONT_FAMILY, 9, "bold"),
    )
    style.map("Treeview.Heading", background=[("active", COLORS["surface_alt"])])
    style.configure(
        "TEntry",
        fieldbackground=COLORS["surface"],
        foreground=COLORS["text"],
        bordercolor=COLORS["border_strong"],
        lightcolor=COLORS["border_strong"],
        darkcolor=COLORS["border_strong"],
        padding=(8, 7),
    )
    style.configure(
        "TCombobox",
        fieldbackground=COLORS["surface"],
        foreground=COLORS["text"],
        bordercolor=COLORS["border_strong"],
        padding=(7, 6),
    )
    style.configure("TNotebook", background=COLORS["surface"], borderwidth=0)
    style.configure("TNotebook.Tab", padding=(14, 8))


def configure_listbox(widget):
    widget.configure(
        background=COLORS["surface"],
        foreground=COLORS["text"],
        selectbackground=COLORS["primary_soft"],
        selectforeground=COLORS["primary"],
        highlightbackground=COLORS["border"],
        highlightcolor=COLORS["primary"],
        highlightthickness=1,
        relief="flat",
        borderwidth=0,
    )


def configure_text(widget, background=None):
    widget.configure(
        background=background or COLORS["surface"],
        foreground=COLORS["text"],
        insertbackground=COLORS["text"],
        highlightbackground=COLORS["border"],
        highlightcolor=COLORS["primary"],
        highlightthickness=1,
        relief="flat",
    )
