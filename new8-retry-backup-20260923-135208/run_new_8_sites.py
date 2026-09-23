# -*- coding: utf-8 -*-
"""Create EIGHT ADDITIONAL sites without overwriting the original eight.

Reuse the tested V2/V3 pipeline and clone, rather than edit, its existing shop-assignment CSV.
The generated sites use .example placeholder domains until real domains are provided.
"""
import csv
import html
import json
import os
import re
import shutil
import subprocess
import sys
import urllib.parse as url
import xml.etree.ElementTree as ET
from collections import defaultdict
from pathlib import Path

FACTORY = Path(r'E:\massage-auto-test')
OUTPUT = Path(r'F:\massage-sites')
SOURCE_ASSIGN = FACTORY / 'factory_v3_shop_assignments_8sites.csv'
ASSIGN = FACTORY / 'factory_v3_shop_assignments_new8.csv'
V2 = FACTORY / 'build_site_factory_v2.py'
V3 = FACTORY / 'build_site_factory_v3.py'
# Old slug -> NEW brand, slug, layout profile, primary colour, background colour.
# The previous eight site directories and assignment file are never altered.
SITE_PAIRS = [
    ('relax-massage', '바로마사지', 'baro-massage', 'A', '#0F766E', '#F0FDFA'),
    ('one-massage',   '애플마사지', 'apple-massage', 'B', '#D63F57', '#FFF5F5'),
    ('two-massage',   '천사마사지', 'cheonsa-massage', 'A', '#2563EB', '#F8FAFC'),
    ('hello-massage', '망고마사지', 'mango-massage', 'B', '#D97706', '#FFFBEB'),
    ('mong-massage',  '바나나마사지', 'banana-massage', 'A', '#CA8A04', '#FEFCE8'),
    ('sarang-massage','하트마사지', 'heart-massage', 'B', '#BE185D', '#FDF2F8'),
    ('new-massage',   '러블리마사지', 'lovely-massage', 'A', '#7C3AED', '#FAF5FF'),
    ('star-massage',  '드림마사지', 'dream-massage', 'B', '#4338CA', '#EEF2FF'),
]
SITES = [pair[1:] for pair in SITE_PAIRS]
ORIGINAL_SLUGS = {pair[0] for pair in SITE_PAIRS}
SLUG_MAP = {pair[0]: pair[2] for pair in SITE_PAIRS}
H1 = re.compile(r'<h1\b[^>]*>(.*?)</h1>', re.I | re.S)
CANON = re.compile(r'<link\b(?=[^>]*\brel\s*=\s*["\']canonical["\'])[^>]*>', re.I | re.S)
HREF = re.compile(r'\bhref\s*=\s*(["\'])(.*?)\1', re.I | re.S)
# Match absolute shop URLs while retaining the path's original encoding convention.
SHOP_URL = re.compile(r'(?P<area>/(?:[^/\s"\'<>?#&]+/)*?)shops/(?P<name>[^/\s"\'<>?#&]+)/')
REL_URL = re.compile(r'(?<![\w/])shops/(?P<name>[^/\s"\'<>?#&]+)/')


def require(cond, message):
    if not cond:
        raise RuntimeError(message)


def get_canonical(text):
    m = CANON.search(text)
    if not m:
        return ''
    h = HREF.search(m.group(0))
    return html.unescape(h.group(2)) if h else ''


def h1_name(text):
    m = H1.search(text)
    return html.unescape(re.sub(r'<[^>]*>', '', m.group(1))).strip() if m else ''


def load_csv(path):
    with path.open('r', encoding='utf-8-sig', newline='') as f:
        return list(csv.DictReader(f))


def check_target(root, slug):
    require(root.resolve() == (OUTPUT / slug).resolve(), '대상 폴더 불일치')
    require(root.is_dir(), '사이트 폴더 없음: ' + str(root))
    require(not root.is_relative_to(FACTORY) and not root.is_relative_to(Path(r'E:\seoul-massage-rebuild')),
            '마스터/운영 사이트 수정 금지')


def update_sitemap(root, canonicals):
    path = root / 'sitemap.xml'
    tree = ET.parse(path)
    xml_root = tree.getroot()
    ns = xml_root.tag.split('}')[0][1:] if xml_root.tag.startswith('{') else ''
    if ns:
        ET.register_namespace('', ns)
    def tag(n):
        return '{' + ns + '}' + n if ns else n
    for child in list(xml_root):
        loc = next((x.text for x in child.iter() if x.tag.rsplit('}', 1)[-1] == 'loc'), None)
        if loc and '/shops/' in url.urlsplit(loc).path:
            xml_root.remove(child)
    old = {x.text for x in xml_root.iter() if x.tag.rsplit('}', 1)[-1] == 'loc'}
    for canonical in dict.fromkeys(canonicals):
        require(canonical not in old, '사이트맵 canonical 중복: ' + canonical)
        n = ET.SubElement(xml_root, tag('url'))
        ET.SubElement(n, tag('loc')).text = canonical
    temp = root / 'sitemap.xml.v4.tmp'
    require(not temp.exists(), '기존 임시 sitemap이 있습니다')
    try:
        tree.write(temp, encoding='utf-8', xml_declaration=True)
        ET.parse(temp)
        os.replace(temp, path)
    finally:
        if temp.exists():
            temp.unlink()


def apply_shop_assignments(root, cfg):
    """Called by a strictly-guarded hook right after V2 generate_content()."""
    root = Path(root).resolve()
    slug = cfg['v4_site_slug']
    check_target(root, slug)
    rows = [r for r in load_csv(Path(cfg['v4_assignments'])) if r['사이트폴더'] == slug]
    require(len(rows) == 5105, f'{slug}: 배정표 5105개가 아닙니다: {len(rows)}')
    by_file = {r['숫자업체페이지']: r for r in rows}
    require(len(by_file) == 5105, '배정표에 중복 업체 페이지가 있습니다')
    local_names = defaultdict(dict)
    for r in rows:
        rel = r['숫자업체페이지']
        require('/shops/' in rel and rel.endswith('/index.html'), '배정표 경로 오류: ' + rel)
        area = rel.split('/shops/', 1)[0]
        require(area == r['지역경로'], '배정표 지역 불일치: ' + rel)
        old, new = r['기존업체명'], r['새업체명']
        require(new and not any(x in new for x in '/\\<>:"|?*') and new[-1] not in ' .',
                'Windows 업체명 경로 부적합: ' + new)
        require(old not in local_names[area] or local_names[area][old] == new,
                '한 지역 기존 업체명이 서로 다르게 배정됨: ' + area)
        local_names[area][old] = new
    for area, values in local_names.items():
        require(len({n.casefold() for n in values.values()}) == len(values),
                '지역 업체명 중복: ' + area)

    numeric = {p.relative_to(root).as_posix(): p for p in root.rglob('index.html')
               if p.parent.name in {'001', '002', '003', '004', '005'}
               and p.parent.parent.name == 'shops'}
    require(set(numeric) == set(by_file), f'숫자 업체 페이지/배정표 불일치: {len(numeric)}개')
    # Preflight ALL files before modifying anything.
    new_canonicals = {}
    old_aliases = {}
    host = None
    for rel, path in numeric.items():
        r = by_file[rel]
        text = path.read_text(encoding='utf-8')
        require(h1_name(text) == r['기존업체명'], 'H1/배정표 불일치: ' + rel)
        old_c = get_canonical(text)
        p = url.urlsplit(old_c)
        require(p.scheme in ('http', 'https') and p.netloc and not p.query and not p.fragment,
                'canonical 형식 오류: ' + rel)
        host = host or p.netloc
        require(p.netloc == host, '혼합 canonical 도메인: ' + rel)
        canonical_file = url.unquote(p.path).lstrip('/') + 'index.html'
        require(canonical_file == r['기존대표페이지'], 'canonical/재고 CSV 불일치: ' + rel)
        if canonical_file == rel:
            new_c = old_c
        else:
            area = r['지역경로']
            old_path = '/' + area + '/shops/' + r['기존업체명'] + '/'
            require(url.unquote(p.path) == old_path, '업체명 canonical 예외: ' + rel)
            new_path = '/' + area + '/shops/' + r['새업체명'] + '/'
            new_c = url.urlunsplit((p.scheme, p.netloc, url.quote(new_path, safe='/'), '', ''))
            old_aliases[(area, r['기존업체명'])] = r['새업체명']
        new_canonicals[rel] = new_c
    require(len(old_aliases) == 2722, f'업체명 대표 URL 개수 확인: {len(old_aliases)}')
    require(len(set(new_canonicals.values())) == 5105, '서로 다른 업체가 canonical URL을 공유합니다')
    print('   배정 사전검사 통과: 5105페이지 / 업체명 URL 2722개', flush=True)

    compiled = {}
    def local_regex(area):
        if area not in compiled:
            keys = sorted(local_names.get(area, {}), key=len, reverse=True)
            compiled[area] = re.compile('|'.join(map(re.escape, keys))) if keys else None
        return compiled[area]

    def transform(text, rel):
        area = rel.removesuffix('/index.html')
        if '/shops/' in area:
            area = area.split('/shops/', 1)[0]
        local = local_names.get(area, {})
        reserved = []
        def protect(value):
            token = '\ue000' + str(len(reserved)) + '\ue001'
            reserved.append(value)
            return token
        def absolute(m):
            raw_area = m.group('area').lstrip('/').rstrip('/')
            a = url.unquote(raw_area)
            old = url.unquote(m.group('name'))
            # Absolute URL regex may also consume the domain as its first segment.
            # Find the longest path suffix that matches an actual region.
            parts = a.split('/')
            new = next((old_aliases[(('/'.join(parts[i:])), old)]
                        for i in range(len(parts))
                        if (('/'.join(parts[i:])), old) in old_aliases), None)
            if new is None:
                return protect(m.group(0)) if local else m.group(0)
            newslug = url.quote(new, safe='') if '%' in m.group('name') else new
            return protect(m.group('area') + 'shops/' + newslug + '/')
        text = SHOP_URL.sub(absolute, text)
        def relative(m):
            old = url.unquote(m.group('name'))
            new = old_aliases.get((area, old))
            if new is None:
                return protect(m.group(0)) if local else m.group(0)
            newslug = url.quote(new, safe='') if '%' in m.group('name') else new
            return protect('shops/' + newslug + '/')
        text = REL_URL.sub(relative, text)
        expr = local_regex(area)
        if expr:
            text = expr.sub(lambda m: local[m.group(0)], text)
        for i, value in enumerate(reserved):
            text = text.replace('\ue000' + str(i) + '\ue001', value)
        return text

    # Only generated site files are modified, never master or live.
    changed = 0
    for path in root.rglob('*.html'):
        rel = path.relative_to(root).as_posix()
        original = path.read_text(encoding='utf-8')
        updated = transform(original, rel)
        if rel in new_canonicals:
            expected = new_canonicals[rel]
            current = get_canonical(updated)
            if current != expected:
                c = CANON.search(updated)
                require(c is not None, 'canonical 태그 없음: ' + rel)
                tag = c.group(0)
                h = HREF.search(tag)
                require(h is not None, 'canonical href 없음: ' + rel)
                new_tag = tag[:h.start(2)] + html.escape(expected, quote=True) + tag[h.end(2):]
                updated = updated[:c.start()] + new_tag + updated[c.end():]
            require(h1_name(updated) == by_file[rel]['새업체명'], '수정 H1 불일치: ' + rel)
            require(get_canonical(updated) == expected, '수정 canonical 불일치: ' + rel)
        if updated != original:
            path.write_text(updated, encoding='utf-8')
            changed += 1
    update_sitemap(root, list(new_canonicals.values()))
    print(f'   지역별 업체명 반영: HTML {changed}개 / canonical 5105개 / sitemap 갱신', flush=True)
    return {'pages': len(rows), 'html_changed': changed}


def install_hook():
    original = V2.read_bytes()
    text = original.decode('utf-8-sig')
    needle = re.compile(r'(?m)^(?P<indent>[ \t]*)content_result\s*=\s*generate_content\(\s*target\s*,\s*content_seed\s*\)')
    found = list(needle.finditer(text))
    require(len(found) == 1, 'V2 업체명 적용 위치를 안전하게 찾지 못했습니다')
    require('apply_shop_assignments(' not in text, 'V2에 업체명 hook이 이미 있습니다')
    indent = found[0].group('indent')
    addition = ('\n' + indent + 'if cfg.get("v4_assignments"):\n'
                + indent + '    from run_new_8_sites import apply_shop_assignments\n'
                + indent + '    apply_shop_assignments(target, cfg)')
    patched = text[:found[0].end()] + addition + text[found[0].end():]
    backup = FACTORY / 'build_site_factory_v2.py.before_new8sites.bak'
    require(not backup.exists(), 'V2 백업이 이미 있습니다. 중복 패치 방지: ' + str(backup))
    backup.write_bytes(original)
    try:
        compile(patched, str(V2), 'exec')
        V2.write_text(patched, encoding='utf-8')
    except BaseException:
        V2.write_bytes(original)
        raise
    return original, backup


def create_config(name, slug, profile, color, background):
    base = FACTORY / ('factory_config_v3_' + profile + '.json')
    cfg = json.loads(base.read_text(encoding='utf-8-sig'))
    require('target' in cfg and 'domain' in cfg and 'brand' in cfg,
            '기존 V3 설정 형식 확인 필요: ' + str(base))
    require(isinstance(cfg['brand'], dict) and 'new' in cfg['brand'],
            'brand.new 설정 구조가 예상과 다릅니다: ' + str(base))
    base_brand = str(cfg['brand']['new'])
    require(bool(base_brand), '기존 브랜드가 비어 있습니다')
    def replace_brand(item):
        if isinstance(item, str):
            return item.replace(base_brand, name)
        if isinstance(item, list):
            return [replace_brand(v) for v in item]
        if isinstance(item, dict):
            return {k: replace_brand(v) for k, v in item.items()}
        return item
    cfg = replace_brand(cfg)
    cfg['target'] = str(OUTPUT / slug)
    former = str(cfg['domain'])
    cfg['domain'] = ('https://' if '://' in former else '') + slug + '.example'
    brand = cfg['brand']
    require(isinstance(brand, dict) and 'new' in brand,
            'brand.new 설정 구조가 예상과 다릅니다: ' + str(base))
    brand['new'] = name
    cfg['content_seed'] = slug + '-new8-' + profile
    cfg['shop_replacements'] = {}
    cfg['v4_assignments'] = str(ASSIGN)
    cfg['v4_site_slug'] = slug
    config_path = FACTORY / ('factory_config_new8site_' + slug + '.json')
    require(not config_path.exists(), '기존 생성 설정을 덮어쓰지 않습니다: ' + str(config_path))
    return config_path, cfg


def prepare_assignments():
    """Read ONLY the original CSV and write a separate NEW-eight-site CSV."""
    rows = load_csv(SOURCE_ASSIGN)
    require(len(rows) == 40840, '원본 배정표가 40840건이 아닙니다: ' + str(len(rows)))
    require({r['사이트폴더'] for r in rows} == ORIGINAL_SLUGS,
            '원본 배정표 사이트 이름이 예상과 다릅니다')
    counts = defaultdict(int)
    for r in rows:
        old = r['사이트폴더']
        counts[old] += 1
        r['사이트폴더'] = SLUG_MAP[old]
    require(all(counts[slug] == 5105 for slug in ORIGINAL_SLUGS),
            '원본 업체 배정 개수가 사이트당 5105건이 아닙니다')
    temp = ASSIGN.with_suffix('.csv.tmp')
    require(not temp.exists(), '기존 임시 배정표가 있습니다: ' + str(temp))
    try:
        with temp.open('w', encoding='utf-8-sig', newline='') as f:
            writer = csv.DictWriter(f, fieldnames=list(rows[0]))
            writer.writeheader()
            writer.writerows(rows)
        os.replace(temp, ASSIGN)
    finally:
        if temp.exists():
            temp.unlink()
    print('새 8개용 배정표 작성 완료 (원본은 보존):', ASSIGN, flush=True)


def main():
    for p in (FACTORY, OUTPUT, SOURCE_ASSIGN, V2, V3):
        require(p.exists(), '필요한 경로/파일 없음: ' + str(p))
    require(Path(__file__).resolve() == (FACTORY / 'run_new_8_sites.py').resolve(),
            '이 스크립트는 E:\\massage-auto-test\\run_new_8_sites.py에 저장해주세요')
    # Full preflight before creating any new files: never overwrite old or new sites/configs.
    for _, slug, *_ in SITES:
        require(not (OUTPUT / slug).exists(), '이미 결과 폴더가 있습니다. 덮어쓰지 않습니다: ' + slug)
        config_path = FACTORY / ('factory_config_new8site_' + slug + '.json')
        require(not config_path.exists(), '이미 config 존재: ' + str(config_path))
    require(not ASSIGN.exists(), '새 배정표가 이미 있습니다. 덮어쓰지 않습니다: ' + str(ASSIGN))
    require(not (FACTORY / 'build_site_factory_v2.py.before_new8sites.bak').exists(),
            '새 작업용 V2 백업이 이미 있습니다. 기존 실행 상태를 확인해주세요')
    prepare_assignments()
    allrows = load_csv(ASSIGN)
    require(len(allrows) == 40840, '8개 사이트 배정표가 40840건이 아닙니다')
    expected = {slug for _, slug, *_ in SITES}
    require({r['사이트폴더'] for r in allrows} == expected, '배정표 사이트 이름 불일치')
    for _, slug, *_ in SITES:
        require(not (OUTPUT / slug).exists(), '이미 결과 폴더가 있습니다. 덮어쓰지 않습니다: ' + slug)
    plans = [create_config(*site) for site in SITES]
    for path, _ in plans:
        require(not path.exists(), '이미 config 존재: ' + str(path))
    for path, cfg in plans:
        path.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding='utf-8')
    original = None
    backup = None
    completed = 0
    try:
        original, backup = install_hook()
        env = os.environ.copy()
        env['PYTHONUNBUFFERED'] = '1'
        for (name, slug, profile, color, background), (config_path, cfg) in zip(SITES, plans):
            print('\n' + '=' * 64 + f'\n[{completed+1}/8] {name} 제작 시작 → {OUTPUT / slug}\n' + '=' * 64, flush=True)
            cmd = [sys.executable, '-u', str(V3), '--config', str(config_path), '--profile', profile]
            result = subprocess.run(cmd, cwd=str(FACTORY), env=env)
            require(result.returncode == 0, f'{name}: V3 생성 중 오류. 나머지 사이트 생성을 중단합니다')
            dest = OUTPUT / slug
            check_target(dest, slug)
            numeric_count = sum(1 for p in dest.rglob('index.html') if p.parent.name in {'001','002','003','004','005'} and p.parent.parent.name == 'shops')
            require(numeric_count == 5105, f'{name}: 결과 숫자 업체 페이지 개수 오류: {numeric_count}')
            html_files = list(dest.rglob('*.html'))
            named_count = sum(1 for p in html_files if p.name == 'index.html'
                              and p.parent.parent.name == 'shops'
                              and p.parent.name not in {'001','002','003','004','005'})
            sitemap_count = sum(1 for e in ET.parse(dest / 'sitemap.xml').getroot().iter()
                                if e.tag.rsplit('}', 1)[-1] == 'loc')
            redirects_text = (dest / '_redirects').read_text(encoding='utf-8-sig')
            marker_a = '# FACTORY CANONICAL 301 START'
            marker_b = '# FACTORY CANONICAL 301 END'
            require(marker_a in redirects_text and marker_b in redirects_text,
                    f'{name}: 301 블록이 없습니다')
            rules = [line for line in redirects_text.split(marker_a, 1)[1].split(marker_b, 1)[0].splitlines()
                     if line.strip() and not line.lstrip().startswith('#')]
            require(len(html_files) == 9460 and named_count == 2722
                    and sitemap_count == 6637 and len(rules) == 2722,
                    f'{name}: 결과 검증 실패 (HTML {len(html_files)}, 업체명 URL {named_count}, '
                    f'sitemap {sitemap_count}, 301 {len(rules)})')
            print(f'   자동 검증 통과: HTML {len(html_files)}, 업체명 URL {named_count}, '
                  f'sitemap {sitemap_count}, 301 {len(rules)}', flush=True)
            # Existing 2 well-tested layout profiles; eight palettes / name distributions.
            from factory_v3_layout import apply_layout
            apply_layout(dest, profile)
            css = dest / 'style.css'
            require(css.is_file(), f'{name}: style.css 없음')
            with css.open('a', encoding='utf-8') as f:
                f.write('\n/* Eight-site palette: ' + slug + ' */\n'
                        + ':root{--factory-brand:' + color + ';--factory-surface:' + background + ';}\n'
                        + 'a:focus-visible,button:focus-visible{outline:3px solid var(--factory-brand);outline-offset:2px}\n')
            completed += 1
            print(f'[{completed}/8] {name} 완료 / 다음 사이트 자동 진행', flush=True)
        print('\n신규 8개 제작 완료. canonical 및 sitemap 도메인은 테스트용 .example입니다. 실제 도메인 확정 전 배포하지 마세요.', flush=True)
    finally:
        if original is not None:
            V2.write_bytes(original)
            if backup is not None:
                print('V2 원본 코드 복원 완료 / 백업 보존:', backup, flush=True)


if __name__ == '__main__':
    try:
        main()
    except Exception as exc:
        print('\n[안전 중단]', exc, file=sys.stderr, flush=True)
        sys.exit(1)
