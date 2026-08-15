"""Review local hymn references vs gdrive_hymn_index.json."""
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from hymn_remote.gdrive_hymn_map import resolve_gdrive_hymn  # noqa: E402

INDEX = ROOT / 'hymn_remote' / 'data' / 'gdrive_hymn_index.json'


def main():
    payload = json.loads(INDEX.read_text(encoding='utf-8'))
    print(f"Drive index: {payload.get('file_count')} files, built {payload.get('built_at')}")

    cases = [
        ('01 神家詩歌', 'P001  001 父的名.doc', '', 'P001  001 父的名.doc'),
        ('05 神家詩歌', 'P02 耶和華尼西', '', ''),
        ('13 神家詩歌', 'P52. 真良人', '', ''),
        ('S1神家詩歌合訂本1大字拍子版', '2\t祢這份愛', 'S1神家詩歌合訂本1大字拍子版 · 2\t祢這份愛 · P.21', ''),
        ('S1神家詩歌合訂本1大字拍子版', '1\t祢這份愛', 'S1神家詩歌合訂本1大字拍子版 · 1\t祢這份愛 · P.18', ''),
        ('S2神家詩歌合訂本2大字拍子版', '10\t超奇愛的神蹟', 'S2神家詩歌合訂本2大字拍子版 · 10\t超奇愛的神蹟 · P.25', ''),
        ('12 神家詩歌', 'P44.永遠愛的傷痕.docx', '', 'P44.永遠愛的傷痕.docx'),
    ]
    ok = 0
    for book, num, label, filename in cases:
        r = resolve_gdrive_hymn(book, num, label, filename=filename)
        hit = 'OK' if r['mapped'] else 'MISS'
        if r['mapped']:
            ok += 1
        print(
            f"[{hit}] {r['code_source']:10} {r['drive_name'][:45]:45} "
            f"| {book} / {num[:25]}"
        )
    print(f'=> {ok}/{len(cases)} mapped')


if __name__ == '__main__':
    main()
