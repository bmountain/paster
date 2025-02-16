import re
import argparse
import json
from pathlib import Path
from .data_model import Config, Argument, Child, Children


def get_number(s: str) -> int:
    """文字列が含む最後の数字を返す"""
    try:
        pattern = r"(\d+)(?!.*\d)"
        return int(re.search(pattern, s).group())
    except:
        raise Exception("画像ファイル名のパースに失敗しました。")


def get_children(dirname: str) -> Children:
    """ディレクトリ名を渡すとその配下のディレクトリとその中の画像一覧を返す"""

    # 引数として与えられたディレクトリの子一覧を取得する
    parent = Path(dirname)
    children_path = [child for child in parent.iterdir() if child.is_dir()]
    try:
        children_path.sort(key=lambda child: int(child.name))
    except:
        raise Exception("子ディレクトリの名前が整数ではありません。")

    children = []
    for child in children_path:
        imgs = list(child.glob("*.png"))
        imgs.sort(key=lambda img: get_number(img.name))
        children.append(Child(**{"path": child, "imgs": imgs}))
    return Children(children)


def parse() -> tuple[str]:
    """引数をパースして画像があるディレクトリ名と出力するファイル名を返す"""
    parser = argparse.ArgumentParser(
        description="指定した親ディレクトリ内の子ディレクトリ配下の画像をエクセルに張り付ける"
    )
    parser.add_argument("--dirname", "-d", required=True, help="親ディレクトリ名")
    parser.add_argument("--out", "-o", required=True, help="出力するエクセルブック名")
    parser.add_argument("--prefix", "-p", required=False, help="項番の接頭辞")
    parser.add_argument("--suffix", "-s", required=False, help="項番の接尾辞")

    args = parser.parse_args()
    dirname, out, prefix, suffix = args.dirname, args.out, args.prefix, args.suffix

    # 拡張子xlsxがなければ補完
    pattern = r".+\.xlsx"
    if not re.fullmatch(pattern, out):
        out += ".xlsx"

    argument = {"dirname": dirname, "out": out, "prefix": prefix, "suffix": suffix}

    return Argument(**argument)


def load_json_config(json_path: Path | None = None) -> Config:
    """jsonを読み込んでConfigを返す"""
    if json_path is None:
        json_path = Path(__file__).parent.parent / Path("config.json")
    with open(json_path, "r") as f:
        config = json.load(f)

    return Config(**config)


def get_config() -> Config:
    """jsonと引数からConfigを作る"""
    config: Config = load_json_config()
    argument: Argument = parse()

    # prefix, suffixは引数を優先する
    if argument.prefix:
        config.prefix = argument.prefix
    if argument.suffix:
        config.suffix = argument.suffix

    config = config.model_dump() | {"dirname": argument.dirname, "out": argument.out}

    return Config(**config)
