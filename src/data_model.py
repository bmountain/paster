from pydantic import BaseModel, Field
from pathlib import Path

# jsonと引数から読み込む設定


class Config(BaseModel):
    """設定全体"""

    # 使用環境
    dpi: int = Field(gt=0, default=96)  # ディスプレイのDPI

    # リサイズ
    max_height: int = Field(gt=0)  # 最大高さ
    max_width: int = Field(gt=0)  # 最大の幅

    # レイアウト
    n_cols: int = Field(gt=0)  # 画像を何列で貼るか
    col_gap: int = Field(ge=0)  # 横並びの画像間で何列空けるか
    row_gap_small: int = Field(ge=0)  # 同じディレクトリ内で改行するとき何行空けるか
    row_gap_big: int = Field(ge=0)  # ディレクトリが切り替わるとき何行空けるか
    row_height_point: int = Field(gt=0)  # Excelシートの行幅

    # 項番
    prefix: str | None  # 項番の接頭辞
    suffix: str | None  # 項番の接尾辞


class Argument(BaseModel):
    """引数"""

    dirname: str  # 親ディレクトリ
    out: str  # 出力するファイル名
    prefix: str | None  # 項番の接頭辞
    suffix: str | None  # 項番の接尾辞


class Params(BaseModel):
    """設定と引数"""

    config: Config
    dirname: str
    out: str


# 画像パスの情報


class Child(BaseModel):
    """pngを含む一つのディレクトリ"""

    path: Path
    imgs: list[Path]


Children = list[Child]
