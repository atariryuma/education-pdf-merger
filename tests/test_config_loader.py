"""
ConfigLoaderのテスト
"""
import os
import pytest
import sys

# プロジェクトルートをパスに追加
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from config_loader import ConfigLoader
from exceptions import ConfigurationError


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
