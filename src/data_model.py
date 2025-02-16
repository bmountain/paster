from pydantic import BaseModel, Field
from pathlib import Path


class Config(BaseModel):
    """設定全体"""

    # 使用環境
    dpi: int = Field(gt=0, default=96, description="ディスプレイのDPI")

    # リサイズ
    max_height: int = Field(gt=0, description="画像の最大高さ（ピクセル）")
    max_width: int = Field(gt=0, description="画像の最大幅（ピクセル）")

    # レイアウト
    n_cols: int = Field(gt=0, description="画像を何列で貼るか")
    col_gap: int = Field(ge=0, description="横並びの画像間で何列空けるか")
    row_gap_small: int = Field(
        ge=0, description="同じディレクトリ内で改行するとき何行空けるか"
    )
    row_gap_big: int = Field(
        ge=0, description="ディレクトリが切り替わるとき何行空けるか"
    )
    row_height_point: int = Field(gt=0, description="Excelシートの行幅")

    # 項番
    prefix: str | None = Field(description="項番の接頭辞")
    suffix: str | None = Field(description="項番の接尾辞")

    dirname: str | None = Field(default=None, description="親ディレクトリ")
    out: str | None = Field(default=None, description="出力するファイルの名前")


class Argument(BaseModel):
    """実行時引数"""

    dirname: str = Field(description="親ディレクトリ")
    out: str = Field(description="出力するファイル名")
    prefix: str | None = Field(default=None, description="項番の接頭辞")
    suffix: str | None = Field(default=None, description="項番の接尾辞")


# 画像パスの情報


class Child(BaseModel):
    """pngを含む一つのディレクトリ"""

    path: Path
    imgs: list[Path]


Children = list[Child]
