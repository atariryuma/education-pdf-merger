"""
ConfigLoaderのテスト
"""
import json
import os
import pytest

from infrastructure.config_loader import ConfigLoader
from shared.exceptions import ConfigurationError


@pytest.mark.unit
class TestConfigLoader:
    """ConfigLoaderクラスのテスト"""

    def test_load_config_success(self, config_file):
        """正常な設定ファイルの読み込み"""
        config = ConfigLoader(config_file)
        assert config.year == "2025"
        assert config.year_short == "R7"

    def test_load_config_file_not_found(self, temp_dir):
        """存在しないファイルの読み込みでConfigurationErrorが発生する"""
        non_existent_path = os.path.join(temp_dir, "non_existent.json")
        with pytest.raises(ConfigurationError) as exc_info:
            ConfigLoader(non_existent_path)
        assert "設定ファイルが見つかりません" in str(exc_info.value)

    def test_load_config_invalid_json(self, temp_dir):
        """不正なJSONの読み込みでConfigurationErrorが発生する"""
        invalid_json_path = os.path.join(temp_dir, "invalid.json")
        with open(invalid_json_path, 'w') as f:
            f.write("{ invalid json }")
        with pytest.raises(ConfigurationError) as exc_info:
            ConfigLoader(invalid_json_path)
        assert "JSON形式が不正" in str(exc_info.value)

    def test_get_nested_keys(self, config_file):
        """ネストされたキーの取得"""
        config = ConfigLoader(config_file)
        assert config.get('base_paths', 'google_drive') == "C:\\TestDrive"
        assert config.get('directories', 'education_plan') == "教育計画書"
        assert config.get('ichitaro', 'max_retries') == 3

    def test_get_with_default(self, config_file):
        """存在しないキーのデフォルト値"""
        config = ConfigLoader(config_file)
        assert config.get('non_existent', default='default_value') == 'default_value'
        assert config.get('base_paths', 'non_existent', default='default') == 'default'

    def test_build_path_with_placeholders(self, config_file):
        """プレースホルダー置換のテスト"""
        config = ConfigLoader(config_file)
        path = config.build_path("{year}", "test")
        assert "2025" in path

        path = config.build_path("{year_short}", "test")
        assert "R7" in path

    def test_get_path_with_dot_notation(self, config_file):
        """ドット記法でのパス取得"""
        config = ConfigLoader(config_file)
        result = config.get_path('base_paths.google_drive', 'test_dir')
        assert "C:\\TestDrive" in result
        assert "test_dir" in result

    def test_get_path_with_validation_success(self, config_file, temp_dir):
        """パス検証（成功）"""
        config = ConfigLoader(config_file)
        # 存在するパスでテスト
        result = config.get_path(temp_dir, validate=True)
        assert result == temp_dir

    def test_get_path_with_validation_failure(self, config_file):
        """パス検証（失敗）"""
        config = ConfigLoader(config_file)
        with pytest.raises(ValueError, match="パスが存在しません"):
            config.get_path("C:\\NonExistent\\Path", validate=True)

    def test_set_and_save(self, config_file, temp_dir):
        """設定の変更と保存"""
        config = ConfigLoader(config_file)
        config.set('test_key', value='test_value')
        assert config.get('test_key') == 'test_value'

        config.set('nested', 'key', value='nested_value')
        assert config.get('nested', 'key') == 'nested_value'

    def test_update_year(self, config_file):
        """年度情報の更新（実運用では西暦のみを使用）"""
        config = ConfigLoader(config_file)
        # GUI での使用方法: 西暦のみ渡す、year_short は自動計算
        config.update_year("2026")
        assert config.year == "2026"
        assert config.year_short == "R8"  # 自動計算される
        assert config.config['year'] == "2026"

    def test_update_year_with_explicit_year_short(self, config_file):
        """年度情報の更新（year_short明示指定）"""
        config = ConfigLoader(config_file)
        # year_shortを明示的に指定した場合
        config.update_year("2027", "R9")
        assert config.year == "2027"
        assert config.year_short == "R9"
        assert config.config['year'] == "2027"
        assert config.config['year_short'] == "R9"

    def test_save_config(self, config_file):
        """設定の保存"""
        config = ConfigLoader(config_file)
        config.set('new_key', value='new_value')
        config.save_config()  # Returns None

        # 再読み込みして確認
        config2 = ConfigLoader(config_file)
        assert config2.get('new_key') == 'new_value'

    def test_get_temp_dir_creates_directory(self, config_file, temp_dir):
        """一時ディレクトリの作成"""
        config = ConfigLoader(config_file)
        # 一時ディレクトリのパスを変更
        new_temp = os.path.join(temp_dir, "new_temp_dir")
        config.set('base_paths', 'local_temp', value=new_temp)

        result = config.get_temp_dir()
        assert result == new_temp
        assert os.path.exists(new_temp)


@pytest.mark.unit
class TestEventNames:
    """get_event_names / save_event_names のテスト"""

    @pytest.fixture
    def config_with_events(self, temp_dir, sample_config_data):
        """デフォルト行事名を含む設定ファイルでConfigLoaderを作成"""
        data = dict(sample_config_data)
        data["excel_default_event_names"] = {
            "school_events": ["入学式", "卒業式", "運動会"],
            "student_council_events": ["生徒総会", "選挙"],
            "other_activities": ["遠足"]
        }
        config_path = os.path.join(temp_dir, "config_events.json")
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        return ConfigLoader(config_path)

    def test_get_event_names_from_default(self, config_with_events):
        """デフォルト行事名が取得できる"""
        names = config_with_events.get_event_names("school_events")
        assert names == ["入学式", "卒業式", "運動会"]

        names = config_with_events.get_event_names("student_council_events")
        assert names == ["生徒総会", "選挙"]

        names = config_with_events.get_event_names("other_activities")
        assert names == ["遠足"]

    def test_get_event_names_returns_copy(self, config_with_events):
        """返されるリストはコピーであり、変更しても内部状態に影響しない"""
        names = config_with_events.get_event_names("school_events")
        original_length = len(names)
        names.append("追加行事")

        # 再取得して変更が影響していないことを確認
        names_again = config_with_events.get_event_names("school_events")
        assert len(names_again) == original_length
        assert "追加行事" not in names_again

    def test_save_and_get_event_names(self, config_with_events):
        """保存した行事名がgetで取得できる"""
        new_names = ["文化祭", "体育祭", "合唱コンクール"]
        config_with_events.save_event_names("school_events", new_names)

        result = config_with_events.get_event_names("school_events")
        assert result == new_names

        # 他のカテゴリには影響しない
        other = config_with_events.get_event_names("student_council_events")
        assert other == ["生徒総会", "選挙"]
