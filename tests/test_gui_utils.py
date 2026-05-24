"""gui.utils のユニットテスト（GUI不要な部分のみ）"""
import tkinter as tk

import pytest


@pytest.fixture
def tk_root():
    """tkinterのルートウィンドウを作成（テスト後に破棄）"""
    try:
        root = tk.Tk()
        root.withdraw()
    except tk.TclError as e:
        pytest.skip(f"tkinter利用不可: {e}")
    yield root
    try:
        root.destroy()
    except tk.TclError:
        pass


@pytest.mark.unit
class TestAttachPlaceholder:
    """attach_placeholder のテスト"""

    def test_empty_var_shows_placeholder(self, tk_root):
        from gui.utils import attach_placeholder
        var = tk.StringVar(value="")
        entry = tk.Entry(tk_root, textvariable=var)
        attach_placeholder(entry, var, "プレースホルダー")
        assert entry.get() == "プレースホルダー"

    def test_existing_value_keeps_value(self, tk_root):
        from gui.utils import attach_placeholder
        var = tk.StringVar(value="既存の値")
        entry = tk.Entry(tk_root, textvariable=var)
        attach_placeholder(entry, var, "プレースホルダー")
        # 既存値があればプレースホルダーは追加されない
        assert entry.get() == "既存の値"

    def test_handlers_registered_for_focus_events(self, tk_root):
        """FocusIn/FocusOutのハンドラがバインドされる"""
        from gui.utils import attach_placeholder
        var = tk.StringVar(value="")
        entry = tk.Entry(tk_root, textvariable=var)
        attach_placeholder(entry, var, "PH")
        # bind()は登録されたbindingの内部名を返す
        assert entry.bind("<FocusIn>") != ""
        assert entry.bind("<FocusOut>") != ""

    def test_placeholder_text_color_is_set(self, tk_root):
        """プレースホルダー表示時に文字色が指定された色になる"""
        from gui.utils import attach_placeholder
        var = tk.StringVar(value="")
        entry = tk.Entry(tk_root, textvariable=var)
        attach_placeholder(entry, var, "PH", color="gray")
        assert str(entry.cget("fg")) == "gray"
