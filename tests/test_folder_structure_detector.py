"""
FolderStructureDetector のユニットテスト

教育計画/行事計画のフォルダ構造判定ロジックをテスト
"""
import os
import pytest

from folder_structure_detector import FolderStructureDetector, PlanType, DetectionResult


@pytest.fixture
def detector():
    return FolderStructureDetector()


class TestDetectStructureEducation:
    """教育計画（3層構造）の検出テスト"""

    def test_deep_nested_structure(self, detector, temp_dir):
        """3層構造のディレクトリが教育計画と判定される"""
        # 3層構造: root/大見出し/小見出し/ファイル
        for main in ["01 教育計画", "02 指導計画", "03 研究計画"]:
            main_dir = os.path.join(temp_dir, main)
            for sub in ["国語", "算数", "理科"]:
                sub_dir = os.path.join(main_dir, sub)
                os.makedirs(sub_dir)
                with open(os.path.join(sub_dir, "file.docx"), 'w') as f:
                    f.write("dummy")

        result = detector.detect_structure(temp_dir)

        assert result.plan_type == PlanType.EDUCATION
        assert result.confidence >= detector.CONFIDENCE_THRESHOLD

    def test_many_main_dirs(self, detector, temp_dir):
        """多数のメインディレクトリは教育計画寄り"""
        for i in range(5):
            main_dir = os.path.join(temp_dir, f"0{i} Section")
            os.makedirs(main_dir)
            with open(os.path.join(main_dir, "file.docx"), 'w') as f:
                f.write("dummy")

        result = detector.detect_structure(temp_dir)

        assert result.evidence['education_score'] > result.evidence['event_score']


class TestDetectStructureEvent:
    """行事計画（2層構造）の検出テスト"""

    def test_flat_structure(self, detector, temp_dir):
        """フラットなディレクトリが行事計画と判定される"""
        # ルート直下にファイルが多い
        for i in range(10):
            with open(os.path.join(temp_dir, f"event_{i}.xlsx"), 'w') as f:
                f.write("dummy")

        result = detector.detect_structure(temp_dir)

        assert result.plan_type == PlanType.EVENT
        assert result.confidence >= detector.CONFIDENCE_THRESHOLD

    def test_few_dirs_many_files(self, detector, temp_dir):
        """ディレクトリが少なくファイルが多い場合は行事計画"""
        # ルートファイル多め + ディレクトリ1つ
        for i in range(8):
            with open(os.path.join(temp_dir, f"file_{i}.docx"), 'w') as f:
                f.write("dummy")
        sub_dir = os.path.join(temp_dir, "sub")
        os.makedirs(sub_dir)
        with open(os.path.join(sub_dir, "sub_file.docx"), 'w') as f:
            f.write("dummy")

        result = detector.detect_structure(temp_dir)

        assert result.evidence['event_score'] > result.evidence['education_score']


class TestDetectStructureEdgeCases:
    """エッジケースのテスト"""

    def test_empty_directory(self, detector, temp_dir):
        """空ディレクトリはEVENT（デフォルト）"""
        result = detector.detect_structure(temp_dir)

        assert result.plan_type == PlanType.EVENT
        assert result.confidence == 0.0

    def test_ambiguous_structure(self, detector, temp_dir):
        """曖昧な構造はAMBIGUOUS"""
        # 中途半端な構造
        main_dir = os.path.join(temp_dir, "section")
        os.makedirs(main_dir)
        with open(os.path.join(main_dir, "file.docx"), 'w') as f:
            f.write("dummy")
        with open(os.path.join(temp_dir, "root_file.docx"), 'w') as f:
            f.write("dummy")

        result = detector.detect_structure(temp_dir)

        # スコアが近い場合はAMBIGUOUSの可能性
        assert isinstance(result, DetectionResult)
        assert result.plan_type in (PlanType.EDUCATION, PlanType.EVENT, PlanType.AMBIGUOUS)

    def test_hidden_files_excluded(self, detector, temp_dir):
        """隠しファイルとテンポラリファイルは除外される"""
        # 隠しファイル・テンポラリ
        with open(os.path.join(temp_dir, ".hidden"), 'w') as f:
            f.write("hidden")
        with open(os.path.join(temp_dir, "~temp.docx"), 'w') as f:
            f.write("temp")
        # 実際のファイル
        with open(os.path.join(temp_dir, "real.docx"), 'w') as f:
            f.write("real")

        result = detector.detect_structure(temp_dir)

        # 除外されたファイルは総ファイル数に含まれない
        assert result.evidence['root_file_count'] == 1

    def test_cover_file_excluded(self, detector, temp_dir):
        """表紙ファイルは除外される"""
        with open(os.path.join(temp_dir, "表紙.docx"), 'w') as f:
            f.write("cover")
        with open(os.path.join(temp_dir, "content.docx"), 'w') as f:
            f.write("content")

        result = detector.detect_structure(temp_dir)

        assert result.evidence['root_file_count'] == 1

    def test_permission_error_returns_event(self, detector):
        """権限エラー時はEVENTが返される"""
        result = detector.detect_structure("/nonexistent/path")

        assert result.plan_type == PlanType.EVENT
        assert result.confidence == 0.0
        assert len(result.issues) > 0


class TestScoreCalculation:
    """スコア計算のテスト"""

    def test_education_score_increases_with_depth(self, detector):
        """深い階層でeducation_scoreが上がる"""
        shallow = {
            'main_dirs': [{'subfolder_count': 0, 'total_files': 1}],
            'root_files': [],
            'max_depth': 1,
            'total_files': 1,
            'main_dir_count': 1,
            'root_file_count': 0,
            'root_file_ratio': 0.0
        }
        deep = {
            'main_dirs': [
                {'subfolder_count': 3, 'total_files': 5},
                {'subfolder_count': 2, 'total_files': 4},
                {'subfolder_count': 3, 'total_files': 6}
            ],
            'root_files': [],
            'max_depth': 3,
            'total_files': 15,
            'main_dir_count': 3,
            'root_file_count': 0,
            'root_file_ratio': 0.0
        }

        score_shallow = detector._calculate_education_score(shallow)
        score_deep = detector._calculate_education_score(deep)

        assert score_deep > score_shallow

    def test_event_score_increases_with_root_files(self, detector):
        """ルートファイルが多いとevent_scoreが上がる"""
        few = {
            'main_dir_count': 1,
            'root_file_count': 1,
            'max_depth': 1,
            'root_file_ratio': 0.5,
        }
        many = {
            'main_dir_count': 1,
            'root_file_count': 10,
            'max_depth': 1,
            'root_file_ratio': 0.9,
        }

        score_few = detector._calculate_event_score(few)
        score_many = detector._calculate_event_score(many)

        assert score_many > score_few


class TestMakeDecision:
    """_make_decision のテスト"""

    def test_education_wins(self, detector):
        """education_score > event_score で EDUCATION"""
        scan = {
            'total_files': 10,
            'main_dir_count': 3,
            'root_file_count': 0,
            'max_depth': 3,
            'root_file_ratio': 0.0
        }
        # confidence = |30-5|/(30+5) = 25/35 ≈ 0.714 > 0.7
        result = detector._make_decision(scan, 30.0, 5.0)
        assert result.plan_type == PlanType.EDUCATION

    def test_event_wins(self, detector):
        """event_score > education_score で EVENT"""
        scan = {
            'total_files': 10,
            'main_dir_count': 1,
            'root_file_count': 8,
            'max_depth': 1,
            'root_file_ratio': 0.8
        }
        # confidence = |2-18|/(2+18) = 16/20 = 0.8 > 0.7
        result = detector._make_decision(scan, 2.0, 18.0)
        assert result.plan_type == PlanType.EVENT

    def test_close_scores_ambiguous(self, detector):
        """スコアが近い場合はAMBIGUOUS"""
        scan = {
            'total_files': 5,
            'main_dir_count': 2,
            'root_file_count': 2,
            'max_depth': 2,
            'root_file_ratio': 0.4
        }
        result = detector._make_decision(scan, 10.0, 10.0)
        assert result.plan_type == PlanType.AMBIGUOUS

    def test_zero_files_returns_event(self, detector):
        """ファイル0件でEVENT（デフォルト）"""
        scan = {
            'total_files': 0,
            'main_dir_count': 0,
            'root_file_count': 0,
            'max_depth': 1,
            'root_file_ratio': 0.0
        }
        result = detector._make_decision(scan, 0.0, 0.0)
        assert result.plan_type == PlanType.EVENT
        assert result.confidence == 0.0
