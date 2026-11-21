"""Build localized Lavender KAG scenarios from the Markdown corpus and 1.xlsx."""
import argparse
import html
import re
import shutil
import zipfile
from pathlib import Path
from xml.etree import ElementTree

TAG = re.compile(r'\[(?:[^\]"\']|"[^"]*"|\'[^\']*\')*\]')
CHOICE = re.compile(r'^(\s*\[seladd\s+text=")([^"]*)(".*)$', re.MULTILINE)
EXTRA_CHOICES = {
    ('lave41.ks', '桜木'): 'Сакураги',
    ('lave41.ks', 'はるか'): 'Харука',
    ('lave41.ks', 'レイカ'): 'Рэйка',
    ('lave41.ks', 'リコ'): 'Рико',
    ('lave41.ks', '沢村'): 'Савамура',
    ('lave.riko2.ks', '目隠しあり'): 'С повязкой на глазах',
    ('lave.riko2.ks', '目隠しなし（画像のみ：テキストの変更はありません）'):
        'Без повязки (меняется только изображение)',
}
# Two Original cells in the corpus already contain Russian.
SOURCE_EXCEPTIONS = {
    ('lave01.ks', 4): 'まずは自己紹介をしようか……。',
    ('lave01.ks', 332): '【ヒカル/？？？】「そうか。ところで、その制服、新山学園の生徒か？」',
}


def read_choices(path):
    def rich_text(node):
        return ''.join(t.text or '' for t in node.findall('.//{*}t'))

    with zipfile.ZipFile(path) as book:
        shared = ([rich_text(item) for item in ElementTree.fromstring(
            book.read('xl/sharedStrings.xml'))]
            if 'xl/sharedStrings.xml' in book.namelist() else [])
        sheet = ElementTree.fromstring(book.read('xl/worksheets/sheet1.xml'))
    choices = {}
    for row in sheet.findall('.//{*}row'):
        fields = {}
        for cell in row.findall('{*}c'):
            value = cell.findtext('{*}v') or ''
            kind = cell.get('t')
            if kind == 's':
                value = shared[int(value)]
            elif kind == 'inlineStr':
                value = rich_text(cell)
            fields[cell.get('r').rstrip('0123456789')] = value.strip()
        if not any(fields.values()) or fields.get('A') == 'ORIGINAL':
            continue
        original, translated, filename = (fields.get(k, '') for k in 'ABC')
        key = filename, original
        if not all((original, translated, filename)) or key in choices:
            raise ValueError(f'{path}: invalid or duplicate choice at row {row.get("r")}')
        choices[key] = translated
    if not choices:
        raise ValueError(f'{path}: no choices')
    return EXTRA_CHOICES | choices


def strip_commands(text):
    """Separate leading inline commands from visible text, keeping [r] as text layout."""
    end = 0
    while (tag := TAG.match(text, end)) and tag[0] != '[r]':
        end = tag.end()
    return text[:end], text[end:]


def markdown_rows(path, scenario):
    rows = []
    lines = path.read_text(encoding='utf-8-sig').splitlines()
    if [value.strip() for value in lines[0].split('|')[1:-1]] != [
        '', 'Original', '', 'TL', 'TLC', 'Edit', 'Notes'
    ]:
        raise ValueError(f'{path}: unexpected Markdown header')
    for number, line in enumerate(lines[2:], 3):
        if not line.strip():
            continue
        cells = re.split(r'(?<!\\)\|', line)
        if len(cells) != 9:
            raise ValueError(f'{path}:{number}: invalid table row')
        source = html.unescape(cells[2].strip()).strip()
        _, source = strip_commands(source)
        if not source:
            continue
        target = html.unescape((cells[5].strip() or cells[4].strip())).strip()
        _, target = strip_commands(target)
        if not target:
            raise ValueError(f'{path}:{number}: missing translation')
        rows.append((number, SOURCE_EXCEPTIONS.get((scenario, number), source),
                     target.replace(r'\|', '|')))
    return rows


def key(text):
    return re.sub(r'\s+', '', text.replace('[r]', r'\n'))


def render_scenario(path, markdown, choices, used_choices):
    original = path.read_bytes().decode('cp932')
    lines = original.splitlines(keepends=True)
    rows = iter(markdown_rows(markdown, path.name))
    output = []
    inserted = 0
    i = 0
    while i < len(lines):
        line = lines[i]
        i += 1
        body = line.rstrip('\r\n')
        newline = line[len(body):]
        stripped = body.strip()
        label, sep, title = body.partition('|')
        if stripped.startswith('*') and sep and title.strip():
            prefix, parts = label + sep, [title]
        elif not stripped or stripped.startswith((';', '*', '@')):
            output.append(line)
            continue
        else:
            prefix, visible = strip_commands(body)
            if not visible.strip():
                output.append(line)
                continue
            # This letter is one Markdown cell but two adjacent physical lines.
            if path.name == 'lave.reika3.ks' and visible == '『これはわたしからの挑戦状よ……':
                continuation = lines[i].rstrip('\r\n')
                if continuation != 'わたしを探したきゃ、探せばいいじゃない……』':
                    raise ValueError(f'{path}:{i+1}: unexpected letter continuation')
                visible += continuation
                i += 1
            parts = visible.split('[*]')
        translated = []
        for source in parts:
            row = next(rows, None)
            if row is None or key(source) != key(row[1]):
                raise ValueError(f'{path}:{i}: source mismatch {source!r}; Markdown {row!r}')
            number, _, target = row
            if target.count(r'\n') != source.count('[r]'):
                raise ValueError(f'{markdown}:{number}: explicit [r] count mismatch')
            translated.append(target.replace(r'\n', '[r]'))
            inserted += 1
        # A timed pause does not provide a separator between Russian sentences.
        joined = translated[0]
        for fragment in translated[1:]:
            previous = joined[-1:]
            following = fragment[:1]
            needs_space = (
                previous and following
                and not previous.isspace() and not following.isspace()
                and not joined.endswith('[r]')
                and (previous in '.!?…」』»' or previous.isalnum())
                and (following.isalnum() or following in '«「『')
            )
            joined += '[*]'
            if needs_space:
                joined += ' '
            joined += fragment
        output.append(prefix + joined + newline)
    if (row := next(rows, None)) is not None:
        raise ValueError(f'{markdown}:{row[0]}: extra text row')

    def insert_choice(match):
        lookup = path.name, match[2].strip()
        if lookup not in choices:
            raise ValueError(f'{path}: missing translation for choice {match[2]!r}')
        used_choices.add(lookup)
        text = re.sub(r'"([^"]*)"', r'«\1»', choices[lookup])
        if '"' in text:
            raise ValueError(f'{path}: unpaired quote in choice {match[2]!r}')
        if '\n' in text or '\r' in text:
            raise ValueError(f'{path}: multiline choice {match[2]!r}')
        return match[1] + text + match[3]

    return CHOICE.sub(insert_choice, ''.join(output)).encode('utf-16'), inserted


def insert(markdown, scenario, output, workbook):
    if output.exists():
        raise ValueError(f'Output directory already exists: {output}')
    choices = read_choices(workbook)
    used_choices = set()
    rendered = {}
    total = 0
    markdown_paths = {p.stem: p for p in markdown.glob('lave*.md')}
    scenarios = sorted(scenario.glob('lave*.ks'))
    if not scenarios:
        raise ValueError(f'No Lavender scenarios in {scenario}')
    for path in scenarios:
        stem = path.stem.replace('lave.', 'lave_', 1)
        if stem not in markdown_paths:
            raise ValueError(f'Missing Markdown for {path.name}')
        data, count = render_scenario(path, markdown_paths.pop(stem), choices, used_choices)
        rendered[path.name] = data
        total += count
    if markdown_paths or choices.keys() - used_choices:
        raise ValueError(f'Unused Markdown files: {list(markdown_paths)}; '
                         f'unmatched choices: {sorted(choices.keys() - used_choices)}')
    shutil.copytree(scenario, output)
    for name, data in rendered.items():
        (output / name).write_bytes(data)
    print(f'{len(rendered)} scenarios, {total} text rows, {len(used_choices)} choices')


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('command', choices=['insert'])
    parser.add_argument('markdown', type=Path)
    parser.add_argument('scenario', type=Path)
    parser.add_argument('output', type=Path)
    parser.add_argument('--choices', type=Path, default=Path(__file__).with_name('1.xlsx'))
    args = parser.parse_args()
    try:
        insert(args.markdown, args.scenario, args.output, args.choices)
    except (OSError, ValueError, zipfile.BadZipFile, ElementTree.ParseError) as error:
        parser.exit(1, f'{error}\n')


if __name__ == '__main__':
    main()
