#!/usr/bin/env python3
"""
ZIP to MSG Converter (템플릿 기반)
실제 Outlook이 생성한 template.msg의 모든 메타데이터/구조를 그대로 사용하고,
첨부 storage의 데이터(파일명/확장자/MIME/바이너리)만 새 zip으로 교체하여
완전히 Outlook 호환되는 .msg 파일을 생성합니다.

같은 폴더의 .zip 파일과 template.msg를 읽어 .msg를 생성합니다.
"""

import struct
import os
import sys
import glob
from datetime import datetime, timezone

ENDOFCHAIN = 0xFFFFFFFE
FREESECT   = 0xFFFFFFFF
NOSTREAM   = 0xFFFFFFFF
FATSECT    = 0xFFFFFFFD
DIFSECT    = 0xFFFFFFFC


def get_exe_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def encode16(s):
    return (s + '\x00').encode('utf-16-le')


def pad512(data):
    r = len(data) % 512
    return data + b'\x00' * (512 - r) if r else data


# ────────────────────────────────────────────────────────────
# OLE2 읽기 (FAT/미니FAT/디렉토리 전부 파싱)
# ────────────────────────────────────────────────────────────

class OleReader:
    def __init__(self, path):
        with open(path, 'rb') as f:
            self.raw = f.read()
        self._parse_header()
        self._parse_fat()
        self._parse_minifat()
        self._parse_dir()

    def _parse_header(self):
        raw = self.raw
        self.dir_start    = struct.unpack_from('<I', raw, 48)[0]
        self.mini_cutoff  = struct.unpack_from('<I', raw, 56)[0]
        self.minifat_start = struct.unpack_from('<I', raw, 60)[0]
        self.minifat_count = struct.unpack_from('<I', raw, 64)[0]
        self.difat_start  = struct.unpack_from('<I', raw, 68)[0]
        self.difat_count  = struct.unpack_from('<I', raw, 72)[0]

    def _sector(self, idx):
        off = 512 + idx * 512
        return self.raw[off:off + 512]

    def _parse_fat(self):
        fat_sectors = []
        for i in range(109):
            v = struct.unpack_from('<I', self.raw, 76 + i * 4)[0]
            if v == FREESECT:
                break
            fat_sectors.append(v)
        cur = self.difat_start
        for _ in range(self.difat_count):
            sec = self._sector(cur)
            entries = struct.unpack_from('<128I', sec, 0)
            fat_sectors.extend([e for e in entries[:127] if e != FREESECT])
            cur = entries[127]
            if cur in (FREESECT, ENDOFCHAIN):
                break
        fat = []
        for s in fat_sectors:
            fat.extend(struct.unpack_from('<128I', self._sector(s), 0))
        self.fat = fat

    def _parse_minifat(self):
        mini_fat = []
        if self.minifat_count and self.minifat_start not in (FREESECT, ENDOFCHAIN):
            cur = self.minifat_start
            seen = 0
            while cur not in (FREESECT, ENDOFCHAIN) and seen < self.minifat_count:
                mini_fat.extend(struct.unpack_from('<128I', self._sector(cur), 0))
                cur = self.fat[cur]
                seen += 1
        self.minifat = mini_fat

    def _chain(self, start):
        if start in (FREESECT, ENDOFCHAIN, NOSTREAM):
            return []
        chain = [start]
        cur = self.fat[start]
        while cur not in (FREESECT, ENDOFCHAIN):
            chain.append(cur)
            cur = self.fat[cur]
        return chain

    def _mini_chain(self, start):
        if start in (FREESECT, ENDOFCHAIN, NOSTREAM):
            return []
        chain = [start]
        cur = self.minifat[start]
        while cur not in (FREESECT, ENDOFCHAIN):
            chain.append(cur)
            cur = self.minifat[cur]
        return chain

    def _parse_dir(self):
        chain = self._chain(self.dir_start)
        entries = []
        for sec in chain:
            data = self._sector(sec)
            for i in range(4):
                entries.append(data[i * 128:(i + 1) * 128])
        self.dir_raw = entries  # 128바이트 원본 그대로 보관 (SID = index)

        # Root Entry의 미니스트림 시작 섹터 확보
        root = entries[0]
        self.ministream_start = struct.unpack_from('<I', root, 116)[0]

    def entry_name(self, sid):
        e = self.dir_raw[sid]
        return e[:64].decode('utf-16-le', errors='replace').rstrip('\x00').strip()

    def entry_fields(self, sid):
        e = self.dir_raw[sid]
        return {
            'etype': e[66],
            'color': e[67],
            'left': struct.unpack_from('<I', e, 68)[0],
            'right': struct.unpack_from('<I', e, 72)[0],
            'child': struct.unpack_from('<I', e, 76)[0],
            'start': struct.unpack_from('<I', e, 116)[0],
            'size': struct.unpack_from('<I', e, 120)[0],
        }

    def read_stream(self, sid):
        f = self.entry_fields(sid)
        size = f['size']
        if size == 0:
            return b''
        if size < self.mini_cutoff:
            mchain = self._mini_chain(f['start'])
            mini_sector_size = 64
            ministream_chain = self._chain(self.ministream_start)
            data = b''
            for sec in ministream_chain:
                data += self._sector(sec)
            out = b''
            for m in mchain:
                off = m * mini_sector_size
                out += data[off:off + mini_sector_size]
            return out[:size]
        else:
            chain = self._chain(f['start'])
            data = b''
            for sec in chain:
                data += self._sector(sec)
            return data[:size]

    def find_child_by_name(self, parent_sid, name):
        """parent_sid의 자식 트리를 순회하며 이름이 일치하는 SID 반환"""
        root_field = self.entry_fields(parent_sid)
        return self._bst_find(root_field['child'], name)

    def _bst_find(self, sid, name):
        if sid == NOSTREAM:
            return None
        ename = self.entry_name(sid)
        if ename == name:
            return sid
        f = self.entry_fields(sid)
        res = self._bst_find(f['left'], name)
        if res is not None:
            return res
        return self._bst_find(f['right'], name)

    def all_children(self, parent_sid):
        """parent_sid 바로 아래 자식 SID 전부 (BST 순회, 정렬됨)"""
        f = self.entry_fields(parent_sid)
        result = []
        self._bst_walk(f['child'], result)
        return result

    def _bst_walk(self, sid, result):
        if sid == NOSTREAM:
            return
        f = self.entry_fields(sid)
        self._bst_walk(f['left'], result)
        result.append(sid)
        self._bst_walk(f['right'], result)


# ────────────────────────────────────────────────────────────
# OLE2 쓰기 (템플릿의 모든 구조를 그대로 복제하며 재조립)
# ────────────────────────────────────────────────────────────

def filetime_now():
    dt = datetime.now(timezone.utc)
    epoch = datetime(1601, 1, 1, tzinfo=timezone.utc)
    return struct.pack('<Q', int((dt - epoch).total_seconds() * 10_000_000))


def build_msg_from_template(template_path, zip_path):
    """
    template_path: 실제 Outlook이 만든 .msg (구조/메타데이터 원본)
    zip_path: 새로 첨부할 zip 파일

    동작:
      1. 템플릿의 모든 스트림(루트+첨부 storage)을 읽어들임
      2. 첨부 storage의 파일명/확장자/MIME/데이터를 새 zip 정보로 교체
      3. 동일한 구조(디렉토리 색상/순서/메타데이터)로 새 OLE2 파일을 재조립
    """
    reader = OleReader(template_path)

    zip_name = os.path.basename(zip_path)
    zip_stem = os.path.splitext(zip_name)[0]
    zip_ext  = os.path.splitext(zip_name)[1]
    mime_type = 'application/x-zip-compressed'
    with open(zip_path, 'rb') as f:
        zip_raw = f.read()

    # 루트(0)와 첨부 storage 찾기
    root_sid = 0
    attach_sid = None
    for sid in range(len(reader.dir_raw)):
        if reader.entry_name(sid).startswith('__attach_version1.0_'):
            attach_sid = sid
            break
    if attach_sid is None:
        raise RuntimeError('템플릿에서 첨부 storage를 찾을 수 없습니다.')

    # 첨부 storage의 자식 스트림들을 이름으로 매핑
    attach_children = reader.all_children(attach_sid)
    attach_stream_by_name = {reader.entry_name(s): s for s in attach_children}

    # 루트의 모든 자식(첨부 storage 포함) 전부 그대로 유지
    root_children = reader.all_children(root_sid)

    # ── 새로 교체할 첨부 관련 스트림 값들 ──
    replacements = {
        '__substg1.0_3001001F': encode16(zip_name),     # display name
        '__substg1.0_3703001F': encode16(zip_ext),      # short filename
        '__substg1.0_3704001F': encode16(zip_ext),      # extension
        '__substg1.0_3707001F': encode16(zip_name),     # long filename
        '__substg1.0_370E001F': encode16(mime_type),    # mime type
        '__substg1.0_37010102': zip_raw,                # 실제 바이너리 데이터
    }

    # ── 모든 스트림 데이터 수집 (SID -> bytes), 교체 대상은 새 값으로 ──
    stream_data = {}   # sid -> bytes (스트림인 경우만)
    storage_sids = set()  # storage(폴더) SID

    def collect(sid):
        f = reader.entry_fields(sid)
        if f['etype'] == 1 or f['etype'] == 5:  # storage or root
            storage_sids.add(sid)
            for child in reader.all_children(sid):
                collect(child)
        elif f['etype'] == 2:  # stream
            name = reader.entry_name(sid)
            if sid in attach_stream_by_name.values() and name in replacements:
                stream_data[sid] = replacements[name]
            else:
                stream_data[sid] = reader.read_stream(sid)

    collect(root_sid)

    # __properties_version1.0 (첨부) 안의 size 필드도 갱신 필요
    # → 해당 스트림은 stream_data에서 직접 재작성
    if '__properties_version1.0' in attach_stream_by_name:
        props_sid = attach_stream_by_name['__properties_version1.0']
        old_props = stream_data[props_sid]
        new_props = bytearray(old_props)
        # 16바이트 엔트리들을 순회하며 가변 길이 항목의 size 갱신
        tag_to_newsize = {
            0x3001: len(replacements['__substg1.0_3001001F']),
            0x3703: len(replacements['__substg1.0_3703001F']),
            0x3704: len(replacements['__substg1.0_3704001F']),
            0x3707: len(replacements['__substg1.0_3707001F']),
            0x370E: len(replacements['__substg1.0_370E001F']),
            0x3701: len(replacements['__substg1.0_37010102']),
        }
        for off in range(8, len(new_props), 16):
            chunk = bytes(new_props[off:off + 16])
            if len(chunk) < 16:
                break
            ptype = struct.unpack_from('<H', chunk, 0)[0]
            tag   = struct.unpack_from('<H', chunk, 2)[0]
            if tag in tag_to_newsize and ptype in (0x001F, 0x0102):
                new_size = tag_to_newsize[tag]
                struct.pack_into('<I', new_props, off + 8, new_size)
        stream_data[props_sid] = bytes(new_props)

    # ── 섹터 할당: 4096바이트 미만 스트림은 미니스트림(64바이트 단위), 이상은 일반 섹터 ──
    MINI_CUTOFF = 4096
    MINI_SECTOR = 64

    def pad64(data):
        r = len(data) % MINI_SECTOR
        return data + b'\x00' * (MINI_SECTOR - r) if r else data

    mini_sectors = []
    mini_fat = []
    sectors = []
    fat = []

    def alloc_big(data):
        if not data:
            return NOSTREAM, 0
        data_padded = pad512(data)
        nsec = len(data_padded) // 512
        start = len(sectors)
        for i in range(nsec):
            sectors.append(data_padded[i * 512:(i + 1) * 512])
            fat.append(start + i + 1 if i < nsec - 1 else ENDOFCHAIN)
        return start, len(data)

    def alloc_mini(data):
        if not data:
            return NOSTREAM, 0
        data_padded = pad64(data)
        nsec = len(data_padded) // MINI_SECTOR
        start = len(mini_sectors)
        for i in range(nsec):
            mini_sectors.append(data_padded[i*MINI_SECTOR:(i+1)*MINI_SECTOR])
            mini_fat.append(start + i + 1 if i < nsec - 1 else ENDOFCHAIN)
        return start, len(data)

    def alloc(data):
        if len(data) < MINI_CUTOFF:
            return alloc_mini(data)
        return alloc_big(data)

    sid_to_loc = {}      # sid -> (start, size)
    sid_is_mini = {}     # sid -> bool
    for sid, data in stream_data.items():
        if len(data) < MINI_CUTOFF:
            sid_to_loc[sid] = alloc_mini(data)
            sid_is_mini[sid] = True
        else:
            sid_to_loc[sid] = alloc_big(data)
            sid_is_mini[sid] = False

    # ── 미니스트림을 일반 섹터에 패킹 ──
    mini_stream_blob = pad512(b''.join(mini_sectors))
    root_ministream_start = NOSTREAM
    root_ministream_size = 0
    if mini_stream_blob:
        nsec = len(mini_stream_blob) // 512
        root_ministream_start = len(sectors)
        for i in range(nsec):
            sectors.append(mini_stream_blob[i*512:(i+1)*512])
            fat.append(root_ministream_start + i + 1 if i < nsec - 1 else ENDOFCHAIN)
        root_ministream_size = len(b''.join(mini_sectors))

    # ── 미니FAT 섹터 (일반 섹터에 저장) ──
    minifat_start_sec = ENDOFCHAIN
    n_minifat_sectors = 0
    if mini_fat:
        rem = len(mini_fat) % 128
        mf = mini_fat + ([FREESECT] * (128 - rem) if rem else [])
        n_minifat_sectors = len(mf) // 128
        minifat_start_sec = len(sectors)
        for i in range(n_minifat_sectors):
            chunk = mf[i*128:(i+1)*128]
            sectors.append(struct.pack('<128I', *chunk))
            fat.append(minifat_start_sec + i + 1 if i < n_minifat_sectors - 1 else ENDOFCHAIN)

    # ── 디렉토리 재작성 (구조/이름/색상/타임스탬프는 원본 유지, start/size만 갱신) ──
    n_entries = len(reader.dir_raw)
    new_dir_entries = [None] * n_entries

    for sid in range(n_entries):
        orig = bytearray(reader.dir_raw[sid])
        f = reader.entry_fields(sid)
        if sid in storage_sids:
            if sid == root_sid:
                # Root Entry: 미니스트림 컨테이너 시작 섹터/크기 기록
                struct.pack_into('<I', orig, 116, root_ministream_start if root_ministream_start != NOSTREAM else ENDOFCHAIN)
                struct.pack_into('<I', orig, 120, root_ministream_size)
            # 일반 storage는 원본 그대로 유지(start=0)
            new_dir_entries[sid] = bytes(orig)
        else:
            # stream: 새로 할당된 위치로 갱신 (미니 스트림이면 미니섹터 인덱스, 아니면 일반 섹터 인덱스)
            start, size = sid_to_loc.get(sid, (NOSTREAM, 0))
            real_start = start if start != NOSTREAM else ENDOFCHAIN
            struct.pack_into('<I', orig, 116, real_start)
            struct.pack_into('<I', orig, 120, size)
            new_dir_entries[sid] = bytes(orig)

    dir_data = pad512(b''.join(new_dir_entries))
    dir_start, _ = alloc_big(dir_data)

    # ── FAT/DIFAT 섹터 계산 ──
    n_data = len(sectors)
    n_fat = max(1, (n_data + 127) // 128)
    n_difat = 0
    for _ in range(5):
        n_difat = max(0, (n_fat - 109 + 126) // 127) if n_fat > 109 else 0
        n_fat = max(1, (n_data + n_fat + n_difat + 127) // 128)

    fat_start = n_data
    difat_start_sec = fat_start + n_fat if n_difat > 0 else ENDOFCHAIN

    fat_full = fat + [FREESECT] * (n_fat * 128 - len(fat))
    for i in range(n_fat):
        fat_full[fat_start + i] = FATSECT
    for i in range(n_difat):
        fat_full[difat_start_sec + i] = DIFSECT

    fat_sectors_data = [struct.pack('<128I', *fat_full[i*128:(i+1)*128]) for i in range(n_fat)]

    fat_refs = list(range(fat_start, fat_start + n_fat))
    header_refs = fat_refs[:109]
    extra_refs  = fat_refs[109:]
    difat_sectors_data = []
    for ci in range(n_difat):
        chunk = extra_refs[ci*127:(ci+1)*127]
        chunk += [FREESECT] * (127 - len(chunk))
        nxt = difat_start_sec + ci + 1 if ci + 1 < n_difat else ENDOFCHAIN
        chunk.append(nxt)
        difat_sectors_data.append(struct.pack('<128I', *chunk))

    # ── 헤더 (미니FAT 사용 안 함 - 전부 일반 섹터로 단순화) ──
    hdr = bytearray(512)
    hdr[0:8] = b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1'
    struct.pack_into('<H', hdr, 24, 0x3E)
    struct.pack_into('<H', hdr, 26, 3)
    struct.pack_into('<H', hdr, 28, 0xFFFE)
    struct.pack_into('<H', hdr, 30, 9)
    struct.pack_into('<H', hdr, 32, 6)
    struct.pack_into('<I', hdr, 40, 0)
    struct.pack_into('<I', hdr, 44, n_fat)
    struct.pack_into('<I', hdr, 48, dir_start)
    struct.pack_into('<I', hdr, 52, 0)
    struct.pack_into('<I', hdr, 56, 0x1000)
    struct.pack_into('<I', hdr, 60, minifat_start_sec)
    struct.pack_into('<I', hdr, 64, n_minifat_sectors)
    struct.pack_into('<I', hdr, 68, difat_start_sec if n_difat > 0 else ENDOFCHAIN)
    struct.pack_into('<I', hdr, 72, n_difat)
    for i, ref in enumerate(header_refs[:109]):
        struct.pack_into('<I', hdr, 76 + i * 4, ref)
    for i in range(len(header_refs), 109):
        struct.pack_into('<I', hdr, 76 + i * 4, FREESECT)

    return (bytes(hdr)
            + b''.join(sectors)
            + b''.join(fat_sectors_data)
            + b''.join(difat_sectors_data))


def main():
    base_dir = get_exe_dir()
    zip_files = glob.glob(os.path.join(base_dir, '*.zip'))
    template_path = os.path.join(base_dir, 'template.msg')

    print('=' * 52)
    print('  ZIP → MSG Converter (템플릿 기반)')
    print('=' * 52)

    if not os.path.exists(template_path):
        print(f'\n[오류] template.msg 파일이 없습니다.')
        print(f'  필요 경로: {template_path}')
        input('\nEnter 키를 눌러 종료...')
        sys.exit(1)

    if not zip_files:
        print(f'\n[오류] 같은 폴더에 .zip 파일이 없습니다.')
        input('\nEnter 키를 눌러 종료...')
        sys.exit(1)

    print(f'\n  {len(zip_files)}개 파일 발견\n')

    ok = fail = 0
    for zip_path in zip_files:
        base = os.path.splitext(zip_path)[0]
        msg_path = base + '.msg'
        fname = os.path.basename(zip_path)
        try:
            out = build_msg_from_template(template_path, zip_path)
            with open(msg_path, 'wb') as f:
                f.write(out)
            print(f'  ✓  {fname}  →  {os.path.basename(msg_path)}')
            ok += 1
        except Exception as e:
            print(f'  ✗  {fname}  ({e})')
            import traceback
            traceback.print_exc()
            fail += 1

    print(f'\n  완료: 성공 {ok}개 / 실패 {fail}개')
    input('\nEnter 키를 눌러 종료...')


if __name__ == '__main__':
    main()
