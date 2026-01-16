"""
フォルダ構造検出モジュール

ディレクトリ構造を分析して教育計画タイプか行事計画タイプかを判定
"""
import logging
import os
from enum import Enum
from dataclasses import dataclass
from typing import Dict, List, Any
from pathlib import Path

logger = logging.getLogger(__name__)


class PlanType(Enum):
    """計画タイプの列挙型"""
    EDUCATION = "education"  # 教育計画（3層構造）
    EVENT = "event"          # 行事計画（2層構造）
    AMBIGUOUS = "ambiguous"  # 判定不能


@dataclass
class DetectionResult:
    """フォルダ構造判定結果"""
    plan_type: PlanType
    confidence: float  # 0.0-1.0
    evidence: Dict[str, Any]
    issues: List[str]


class FolderStructureDetector:
    """フォルダ構造検出器（パフォーマンス最適化版）"""

    # 設定値
    CONFIDENCE_THRESHOLD = 0.7
    SCORE_DIFFERENCE_THRESHOLD = 3.0
    MIN_MAIN_DIRS_FOR_EDUCATION = 2
    MAX_ROOT_FILE_RATIO_FOR_EDUCATION = 0.3

    # パフォーマンス設定
    MAX_SCAN_DEPTH = 10  # 最大スキャン深度（深いプロジェクトにも対応）

    def detect_structure(self, directory_path: str) -> DetectionResult:
        """
        ディレクトリ構造を分析して計画タイプを判定

        Args:
            directory_path: 分析対象のディレクトリパス

        Returns:
            DetectionResult: 判定結果
        """
        logger.info(f"フォルダ構造分析を開始: {directory_path}")

        try:
            # ステップ1: 構造をスキャン
            scan_result = self._scan_directory(directory_path)

            # ステップ2: スコアリング
            education_score = self._calculate_education_score(scan_result)
            event_score = self._calculate_event_score(scan_result)

            # ステップ3: 判定
            result = self._make_decision(
                scan_result, education_score, event_score
            )

            logger.info(
                f"判定完了: {result.plan_type.value} "
                f"(確信度: {result.confidence:.2f}, "
                f"教育={education_score:.1f}, 行事={event_score:.1f})"
            )

            return result

        except Exception as e:
            logger.error(f"フォルダ構造分析エラー: {e}", exc_info=True)
            # エラー時は行事計画（デフォルト）
            return DetectionResult(
                plan_type=PlanType.EVENT,
                confidence=0.0,
                evidence={"error": str(e)},
                issues=[f"分析エラー: {e}"]
            )

    def _scan_directory(self, directory_path: str) -> Dict[str, Any]:
        """ディレクトリ構造をスキャン"""
        path = Path(directory_path)

        # 除外パターン（表紙ファイル、隠しファイル等）
        def should_exclude(name: str) -> bool:
            return (
                name.startswith('.') or
                name.startswith('~') or
                name == '__pycache__' or
                '表紙' in name
            )

        try:
            items = [
                item for item in os.listdir(directory_path)
                if not should_exclude(item)
            ]
        except PermissionError as e:
            logger.error(f"ディレクトリアクセス権限エラー: {e}")
            raise

        main_dirs = []
        root_files = []
        max_depth = 1

        for item in items:
            item_path = path / item

            if item_path.is_dir():
                # メインディレクトリの分析
                dir_info = self._analyze_directory(item_path, depth=2)
                main_dirs.append(dir_info)
                max_depth = max(max_depth, dir_info['max_depth'])
            elif item_path.is_file():
                root_files.append(item)

        total_files = len(root_files) + sum(
            d['total_files'] for d in main_dirs
        )

        result = {
            'main_dirs': main_dirs,
            'root_files': root_files,
            'max_depth': max_depth,
            'total_files': total_files,
            'main_dir_count': len(main_dirs),
            'root_file_count': len(root_files),
            'root_file_ratio': (
                len(root_files) / total_files if total_files > 0 else 0
            )
        }

        logger.debug(
            f"スキャン結果: メインディレクトリ={len(main_dirs)}個, "
            f"ルートファイル={len(root_files)}個, "
            f"最大深度={max_depth}層, "
            f"総ファイル数={total_files}個"
        )

        return result

    def _analyze_directory(
        self, dir_path: Path, depth: int
    ) -> Dict[str, Any]:
        """
        ディレクトリを再帰的に分析

        Args:
            dir_path: 分析対象のディレクトリPath
            depth: 現在の階層深度

        Returns:
            Dict[str, Any]: ディレクトリ情報
        """
        # 深さ制限
        if depth > self.MAX_SCAN_DEPTH:
            return {
                'subfolders': [],
                'files': [],
                'max_depth': depth,
                'subfolder_count': 0,
                'file_count': 0,
                'total_files': 0
            }

        subfolders = []
        files = []
        max_depth = depth

        try:
            # iteratorを直接使用（メモリ効率改善）
            items_iter = dir_path.iterdir()
        except PermissionError:
            logger.warning(f"アクセス権限なし: {dir_path}")
            return {
                'subfolders': [],
                'files': [],
                'max_depth': depth,
                'subfolder_count': 0,
                'file_count': 0,
                'total_files': 0
            }

        # os.scandir()を使用することで、is_symlink()と属性取得が1回のシステムコールで済む
        # ctypesを使わずPythonネイティブのis_symlink()のみ使用（パフォーマンス大幅改善）
        for item in items_iter:
            if item.name.startswith('.') or item.name.startswith('~'):
                continue

            try:
                # シンボリックリンク・ジャンクション・リパースポイント検出
                # os.scandir()のis_symlink()はWindowsジャンクションも検出可能
                if item.is_symlink():
                    logger.debug(f"シンボリックリンク/ジャンクションをスキップ: {item}")
                    continue

                if item.is_dir():
                    sub_info = self._analyze_directory(item, depth + 1)
                    subfolders.append(sub_info)
                    max_depth = max(max_depth, sub_info['max_depth'])
                elif item.is_file():
                    files.append(item.name)
            except OSError as e:
                # シンボリックリンクエラー、権限エラー等をスキップ
                logger.debug(f"ファイル/フォルダアクセスエラー（スキップ）: {item}, エラー: {e}")
                continue

        total_files = len(files) + sum(
            s['total_files'] for s in subfolders
        )

        return {
            'subfolders': subfolders,
            'files': files,
            'max_depth': max_depth,
            'subfolder_count': len(subfolders),
            'file_count': len(files),
            'total_files': total_files
        }

    def _calculate_education_score(
        self, scan_result: Dict[str, Any]
    ) -> float:
        """教育計画（3層構造）スコアを計算"""
        score = 0.0

        # メインディレクトリ数（多いほど高スコア）
        main_dir_count = scan_result['main_dir_count']
        score += main_dir_count * 2.0

        # サブフォルダの平均数
        if main_dir_count > 0:
            avg_subfolders = sum(
                d['subfolder_count'] for d in scan_result['main_dirs']
            ) / main_dir_count
            score += avg_subfolders * 1.5

        # 階層の深さ
        if scan_result['max_depth'] >= 3:
            score += 3.0

        # ルートファイル比率が低い
        if scan_result['root_file_ratio'] < self.MAX_ROOT_FILE_RATIO_FOR_EDUCATION:
            score += 2.0

        return score

    def _calculate_event_score(self, scan_result: Dict[str, Any]) -> float:
        """行事計画（2層構造）スコアを計算"""
        score = 0.0

        # ルートファイル数（多いほど高スコア）
        score += scan_result['root_file_count'] * 1.0

        # サブフォルダがない／少ない
        if scan_result['main_dir_count'] <= 3:
            score += 1.5

        # 階層が浅い
        if scan_result['max_depth'] <= 2:
            score += 3.0

        # ルートファイル比率が高い
        if scan_result['root_file_ratio'] > 0.5:
            score += 2.0

        return score

    def _make_decision(
        self,
        scan_result: Dict[str, Any],
        education_score: float,
        event_score: float
    ) -> DetectionResult:
        """スコアから最終判定"""
        score_diff = abs(education_score - event_score)
        max_score = max(education_score, event_score)
        total_score = education_score + event_score

        # 確信度の計算
        if total_score > 0:
            confidence = score_diff / total_score
        else:
            confidence = 0.0

        # 証拠データ
        evidence = {
            'education_score': education_score,
            'event_score': event_score,
            'score_difference': score_diff,
            'main_dir_count': scan_result['main_dir_count'],
            'root_file_count': scan_result['root_file_count'],
            'max_depth': scan_result['max_depth'],
            'root_file_ratio': scan_result['root_file_ratio']
        }

        # 判定
        issues = []

        # 空ディレクトリの処理
        if scan_result['total_files'] == 0:
            issues.append("ディレクトリにファイルが存在しません")
            return DetectionResult(
                plan_type=PlanType.EVENT,  # デフォルト
                confidence=0.0,
                evidence=evidence,
                issues=issues
            )

        # 確信度による判定
        if confidence < self.CONFIDENCE_THRESHOLD:
            plan_type = PlanType.AMBIGUOUS
            issues.append("判定の確信度が低いため、手動選択を推奨します")
            logger.info(
                f"判定が曖昧: 確信度={confidence:.2f} < {self.CONFIDENCE_THRESHOLD}"
            )
        elif education_score > event_score:
            plan_type = PlanType.EDUCATION
            logger.info(
                f"教育計画と判定: スコア差={score_diff:.1f}, "
                f"確信度={confidence:.2f}"
            )
        else:
            plan_type = PlanType.EVENT
            logger.info(
                f"行事計画と判定: スコア差={score_diff:.1f}, "
                f"確信度={confidence:.2f}"
            )

        return DetectionResult(
            plan_type=plan_type,
            confidence=confidence,
            evidence=evidence,
            issues=issues
        )
