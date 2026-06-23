#!/usr/bin/env python3
"""
ZIP to MSG Converter
같은 폴더의 .zip 파일을 전부 첨부파일이 포함된 .msg 파일로 변환합니다.
실제 Outlook MSG 구조(Mini FAT 포함)를 따릅니다.
"""

import struct, os, sys, glob
from datetime import datetime, timezone

ENDOFCHAIN   = 0xFFFFFFFE
FREESECT     = 0xFFFFFFFF
NOSTREAM     = 0xFFFFFFFF
FATSECT      = 0xFFFFFFFD
DIFSECT      = 0xFFFFFFFC
MINI_CUTOFF  = 4096   # 이 크기 미만 스트림은 미니 스트림으로 저장
MINI_SECTOR  = 64     # 미니 섹터 크기


def get_exe_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def encode16(s):
    return (s + '\x00').encode('utf-16-le')


def pad_to(data, n):
    r = len(data) % n
    return data + b'\x00' * (n - r) if r else data


def pad512(d):
    return pad_to(d, 512)


def pad64(d):
    return pad_to(d, MINI_SECTOR)


def filetime_now():
    dt = datetime.now(timezone.utc)
    epoch = datetime(1601, 1, 1, tzinfo=timezone.utc)
    return struct.pack('<Q', int((dt - epoch).total_seconds() * 10_000_000))


def de(name, etype, start, size,
       child=NOSTREAM, left=NOSTREAM, right=NOSTREAM, color=1):
    """OLE2 디렉토리 엔트리 128바이트"""
    ne   = name.encode('utf-16-le')[:62]
    nlen = len(ne) + 2 if ne else 0
    ne   = ne.ljust(64, b'\x00')
    b    = struct.pack('<H', nlen) + struct.pack('<B', etype) + struct.pack('<B', color)
    b   += struct.pack('<III', left, right, child)
    b   += b'\x00' * 16 + struct.pack('<IQQ', 0, 0, 0)
    b   += struct.pack('<II', start if start != NOSTREAM else ENDOFCHAIN, size)
    b   += b'\x00' * 4
    assert len(ne) + len(b) == 128
    return ne + b


def empty_de():
    """패딩용 빈 디렉토리 엔트리"""
    return b'\x00' * 64 + struct.pack('<HBBI', 0, 0, 1, NOSTREAM) + \
           struct.pack('<III', NOSTREAM, NOSTREAM, NOSTREAM) + \
           b'\x00' * 16 + struct.pack('<IQQ', 0, 0, 0) + \
           struct.pack('<II', ENDOFCHAIN, 0) + b'\x00' * 4


def ole_sort_key(name):
    """OLE2 디렉토리 정렬 규칙: 이름 길이 우선, 그다음 대소문자 무시 비교"""
    return (len(name), name.upper())


def build_storage_chain(children, sid_offset):
    """
    같은 부모를 가진 자식 엔트리들을 OLE2 정렬 규칙에 따라 정렬한 뒤
    단순 사슬(각 노드의 right만 다음 노드를 가리킴, 퇴화 BST) 형태의
    디렉토리 엔트리 리스트를 만든다. 모든 색은 black(1)으로 통일.

    children: [(name, etype, start, size, child_sid_or_None)]
    sid_offset: 이 그룹의 첫 엔트리가 배치될 SID
    반환: (정렬된 dir_entry 바이트 리스트, 첫 엔트리의 SID)
    """
    ordered = sorted(children, key=lambda c: ole_sort_key(c[0]))
    entries = []
    n = len(ordered)
    for i, (name, etype, start, size, child_sid) in enumerate(ordered):
        right = sid_offset + i + 1 if i + 1 < n else NOSTREAM
        child = child_sid if child_sid is not None else NOSTREAM
        entries.append(de(name, etype, start, size, child=child, right=right, color=1))
    first_sid = sid_offset if n else NOSTREAM
    return entries, first_sid


def build_root_props(entries, next_recipient_id=0, next_attachment_id=1,
                      recipient_count=0, attachment_count=1):
    """
    루트 properties 헤더 32바이트 (MS-OXMSG):
      reserved(8) + nextRecipientId(4) + nextAttachmentId(4)
      + recipientCount(4) + attachmentCount(4) + reserved(8)
    엔트리는 16바이트: ptype(2)+tag(2)+flag(4)+value(8)
    """
    hdr = b'\x00' * 8
    hdr += struct.pack('<I', next_recipient_id)
    hdr += struct.pack('<I', next_attachment_id)
    hdr += struct.pack('<I', recipient_count)
    hdr += struct.pack('<I', attachment_count)
    hdr += b'\x00' * 8
    body = b''.join(entries)
    return hdr + body


def build_att_props(entries):
    """첨부 properties: 헤더 8바이트(reserved) + 16바이트 엔트리 반복"""
    body = b''.join(entries)
    return b'\x00' * 8 + body


def prop_fixed(ptype, tag, value8):
    """고정 크기 프로퍼티 16바이트: ptype+tag+flag(4)+value(8, 부족하면 0패딩)
    실제 Outlook 파일 분석 결과 flag는 고정/가변 무관하게 항상 0x00000006."""
    value8 = value8[:8].ljust(8, b'\x00')
    return struct.pack('<HH', ptype, tag) + struct.pack('<I', 6) + value8


def prop_var(ptype, tag, size):
    """가변 크기 프로퍼티 16바이트: ptype+tag+flag(6)+size(4)+reserved(3)"""
    return struct.pack('<HH', ptype, tag) + struct.pack('<I', 6) + struct.pack('<I', size) + struct.pack('<I', 3)


class MsgBuilder:
    """
    OLE2 Compound File 빌더.
    - 4096바이트 이상 스트림 → 일반 섹터(512바이트) + 일반 FAT
    - 4096바이트 미만 스트림 → 미니 스트림(64바이트 섹터) + 미니 FAT
    """

    def __init__(self):
        self.big_sectors  = []   # 512바이트 일반 섹터들
        self.big_fat      = []
        self.mini_sectors = []   # 64바이트 미니 섹터들 (나중에 512배수로 묶어 일반 섹터에 저장)
        self.mini_fat      = []
        self.streams = []        # (key, raw_size, is_mini, start)

    def add_stream(self, key, data):
        """실제 데이터(패딩 전)를 받아 적절히 배치, 실제 크기를 size로 사용"""
        raw_size = len(data)
        if raw_size == 0:
            self.streams.append((key, 0, False, NOSTREAM))
            return
        if raw_size < MINI_CUTOFF:
            # 미니 스트림에 64바이트 단위로 저장
            padded = pad64(data)
            nsec = len(padded) // MINI_SECTOR
            start = len(self.mini_sectors)
            for i in range(nsec):
                self.mini_sectors.append(padded[i*MINI_SECTOR:(i+1)*MINI_SECTOR])
                self.mini_fat.append(start + i + 1 if i < nsec - 1 else ENDOFCHAIN)
            self.streams.append((key, raw_size, True, start))
        else:
            # 일반 섹터에 512바이트 단위로 저장
            padded = pad512(data)
            nsec = len(padded) // 512
            start = len(self.big_sectors)
            for i in range(nsec):
                self.big_sectors.append(padded[i*512:(i+1)*512])
                self.big_fat.append(start + i + 1 if i < nsec - 1 else ENDOFCHAIN)
            self.streams.append((key, raw_size, False, start))

    def info(self, key):
        """(start, size, is_mini) 반환"""
        for k, sz, is_mini, start in self.streams:
            if k == key:
                return start, sz, is_mini
        raise KeyError(key)

    def finalize(self, dir_entries_builder):
        """
        dir_entries_builder(info_func) -> list[dir_entry_bytes]
        디렉토리 엔트리들을 생성하고 전체 OLE2 파일 바이트 반환
        """
        # 1) 미니스트림을 512바이트 단위로 패킹하여 일반 섹터에 저장
        mini_stream_blob = pad512(b''.join(self.mini_sectors))
        root_start = NOSTREAM
        root_size  = 0
        if mini_stream_blob:
            nsec = len(mini_stream_blob) // 512
            root_start = len(self.big_sectors)
            for i in range(nsec):
                self.big_sectors.append(mini_stream_blob[i*512:(i+1)*512])
                self.big_fat.append(root_start + i + 1 if i < nsec - 1 else ENDOFCHAIN)
            root_size = len(b''.join(self.mini_sectors))  # 실제 미니스트림 크기

        # 2) 미니 FAT 섹터 (128개씩 일반 섹터에 저장)
        minifat_start = ENDOFCHAIN
        n_minifat_sectors = 0
        if self.mini_fat:
            rem = len(self.mini_fat) % 128
            mf  = self.mini_fat + ([FREESECT] * (128 - rem) if rem else [])
            n_minifat_sectors = len(mf) // 128
            minifat_start = len(self.big_sectors)
            for i in range(n_minifat_sectors):
                chunk = mf[i*128:(i+1)*128]
                self.big_sectors.append(struct.pack('<128I', *chunk))
                self.big_fat.append(minifat_start + i + 1 if i < n_minifat_sectors - 1 else ENDOFCHAIN)

        # 3) 디렉토리 엔트리 생성 (콜백으로 위임)
        def info_func(key):
            start, size, is_mini = self.info(key)
            return start, size
        dir_entries = dir_entries_builder(info_func, root_start, root_size)
        while len(dir_entries) % 4:
            dir_entries.append(empty_de())
        dir_data = pad512(b''.join(dir_entries))
        dir_nsec = len(dir_data) // 512
        dir_start = len(self.big_sectors)
        for i in range(dir_nsec):
            self.big_sectors.append(dir_data[i*512:(i+1)*512])
            self.big_fat.append(dir_start + i + 1 if i < dir_nsec - 1 else ENDOFCHAIN)

        # 4) FAT + DIFAT 섹터 계산
        n_data = len(self.big_sectors)
        n_fat = max(1, (n_data + 127) // 128)
        n_difat = 0
        for _ in range(5):
            n_difat = max(0, (n_fat - 109 + 126) // 127) if n_fat > 109 else 0
            n_fat = max(1, (n_data + n_fat + n_difat + 127) // 128)

        fat_start = n_data
        difat_start = fat_start + n_fat if n_difat > 0 else ENDOFCHAIN

        fat_full = self.big_fat + [FREESECT] * (n_fat * 128 - len(self.big_fat))
        for i in range(n_fat):
            fat_full[fat_start + i] = FATSECT
        for i in range(n_difat):
            fat_full[difat_start + i] = DIFSECT

        fat_sectors_data = [struct.pack('<128I', *fat_full[i*128:(i+1)*128]) for i in range(n_fat)]

        fat_refs = list(range(fat_start, fat_start + n_fat))
        header_refs = fat_refs[:109]
        extra_refs  = fat_refs[109:]
        difat_sectors_data = []
        for ci in range(n_difat):
            chunk = extra_refs[ci*127:(ci+1)*127]
            chunk += [FREESECT] * (127 - len(chunk))
            nxt = difat_start + ci + 1 if ci + 1 < n_difat else ENDOFCHAIN
            chunk.append(nxt)
            difat_sectors_data.append(struct.pack('<128I', *chunk))

        # 5) 헤더
        hdr = bytearray(512)
        hdr[0:8] = b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1'
        struct.pack_into('<H', hdr, 24, 0x3E)
        struct.pack_into('<H', hdr, 26, 3)
        struct.pack_into('<H', hdr, 28, 0xFFFE)
        struct.pack_into('<H', hdr, 30, 9)   # sector size exp -> 512
        struct.pack_into('<H', hdr, 32, 6)   # mini sector size exp -> 64
        struct.pack_into('<I', hdr, 40, 0)
        struct.pack_into('<I', hdr, 44, n_fat)
        struct.pack_into('<I', hdr, 48, dir_start)
        struct.pack_into('<I', hdr, 52, 0)
        struct.pack_into('<I', hdr, 56, MINI_CUTOFF)
        struct.pack_into('<I', hdr, 60, minifat_start)
        struct.pack_into('<I', hdr, 64, n_minifat_sectors)
        struct.pack_into('<I', hdr, 68, difat_start if n_difat > 0 else ENDOFCHAIN)
        struct.pack_into('<I', hdr, 72, n_difat)
        for i, ref in enumerate(header_refs[:109]):
            struct.pack_into('<I', hdr, 76 + i*4, ref)
        for i in range(len(header_refs), 109):
            struct.pack_into('<I', hdr, 76 + i*4, FREESECT)

        return (bytes(hdr)
                + b''.join(self.big_sectors)
                + b''.join(fat_sectors_data)
                + b''.join(difat_sectors_data))


def build_zip_msg(zip_path):
    zip_name  = os.path.basename(zip_path)
    zip_stem  = os.path.splitext(zip_name)[0]
    zip_ext   = os.path.splitext(zip_name)[1]
    mime_type = 'application/x-zip-compressed'

    with open(zip_path, 'rb') as f:
        zip_raw = f.read()

    sender_name  = zip_stem
    sender_email = f'{zip_stem}@attachment.msg'
    conv_topic   = zip_stem
    body_text    = f'Attachment: {zip_name}'

    b = MsgBuilder()

    # ── 루트 스트림 데이터 ──
    mc      = encode16('IPM.Note')
    subj    = encode16(zip_name)
    sname   = encode16(sender_name)
    semail  = encode16(sender_email)
    stype   = encode16('SMTP')
    ctopic  = encode16(conv_topic)
    body_b  = encode16(body_text)
    dispto  = encode16('')

    root_prop = build_root_props([
        prop_fixed(0x0040, 0x0039, filetime_now()),
        prop_var(0x001F, 0x001A, len(mc)),
        prop_var(0x001F, 0x0037, len(subj)),
        prop_var(0x001F, 0x0070, len(ctopic)),
        prop_var(0x001F, 0x0C1A, len(sname)),
        prop_var(0x001F, 0x0C1F, len(semail)),
        prop_var(0x001F, 0x0C1E, len(stype)),
        prop_var(0x001F, 0x0042, len(sname)),
        prop_var(0x001F, 0x0065, len(semail)),
        prop_var(0x001F, 0x0064, len(stype)),
        prop_var(0x001F, 0x0E04, len(dispto)),
        prop_var(0x001F, 0x1000, len(body_b)),
        prop_fixed(0x000B, 0x0E1B, struct.pack('<I', 1)),
    ])

    b.add_stream('root_prop', root_prop)
    b.add_stream('mc', mc)
    b.add_stream('subj', subj)
    b.add_stream('sname', sname)
    b.add_stream('semail', semail)
    b.add_stream('stype', stype)
    b.add_stream('ctopic', ctopic)
    b.add_stream('body', body_b)
    b.add_stream('dispto', dispto)

    # ── 첨부 스트림 데이터 ──
    ext       = encode16(zip_ext)
    lname     = encode16(zip_name)       # 긴 파일명 (3707, 3001)
    shortname = encode16(zip_ext)        # 단축 파일명은 확장자만 (3703) — 실제 Outlook 동작과 동일
    mime      = encode16(mime_type)
    locale    = encode16('EnUs')

    att_prop = build_att_props([
        prop_fixed(0x0003, 0x0E21, struct.pack('<I', 0)),
        prop_fixed(0x0003, 0x3705, struct.pack('<I', 1)),
        prop_var(0x001F, 0x3001, len(lname)),
        prop_var(0x001F, 0x3703, len(shortname)),
        prop_var(0x001F, 0x3704, len(ext)),
        prop_fixed(0x0003, 0x370B, struct.pack('<i', -1)),  # PR_RENDERING_POSITION = -1 (렌더링 안 함)
        prop_var(0x001F, 0x3707, len(lname)),
        prop_var(0x001F, 0x370E, len(mime)),
        prop_var(0x001F, 0x3A0C, len(locale)),
        prop_var(0x0102, 0x3701, len(zip_raw)),
    ])

    b.add_stream('att_prop', att_prop)
    b.add_stream('ext', ext)
    b.add_stream('lname', lname)
    b.add_stream('shortname', shortname)
    b.add_stream('mime', mime)
    b.add_stream('locale', locale)
    b.add_stream('zipdata', zip_raw)   # 큰 데이터는 자동으로 일반 섹터 사용

    def build_dirs(info, root_start, root_size):
        # ── 첨부 스토리지의 자식들 (정렬 후 사슬 구성) ──
        att_children = [
            ('__properties_version1.0', 2, *info('att_prop'), None),
            ('__substg1.0_3704001F',     2, *info('ext'),       None),
            ('__substg1.0_3703001F',     2, *info('shortname'), None),
            ('__substg1.0_3707001F',     2, *info('lname'),     None),
            ('__substg1.0_3001001F',     2, *info('lname'),     None),
            ('__substg1.0_370E001F',     2, *info('mime'),      None),
            ('__substg1.0_3A0C001F',     2, *info('locale'),    None),
            ('__substg1.0_37010102',     2, *info('zipdata'),   None),
        ]
        # SID 배치: 0=Root, 1..9=루트 자식(9개), 10=첨부 스토리지, 11..18=첨부 자식(8개)
        att_entries, att_first = build_storage_chain(att_children, sid_offset=11)

        # ── 루트의 자식들 (첨부 스토리지 포함) ──
        root_children = [
            ('__properties_version1.0',      2, *info('root_prop'), None),
            ('__substg1.0_001A001F',          2, *info('mc'),       None),
            ('__substg1.0_0037001F',          2, *info('subj'),     None),
            ('__substg1.0_0C1A001F',          2, *info('sname'),    None),
            ('__substg1.0_0C1F001F',          2, *info('semail'),   None),
            ('__substg1.0_0C1E001F',          2, *info('stype'),    None),
            ('__substg1.0_0070001F',          2, *info('ctopic'),   None),
            ('__substg1.0_1000001F',          2, *info('body'),     None),
            ('__substg1.0_0E04001F',          2, *info('dispto'),   None),
            ('__attach_version1.0_#00000000', 1, NOSTREAM, 0,        att_first),
        ]
        root_entries, root_first = build_storage_chain(root_children, sid_offset=1)

        dirs = [de('Root Entry', 5, root_start, root_size, child=root_first, color=0)]
        dirs.extend(root_entries)
        dirs.extend(att_entries)
        return dirs

    return b.finalize(build_dirs)


def main():
    base_dir  = get_exe_dir()
    zip_files = glob.glob(os.path.join(base_dir, '*.zip'))

    print('=' * 52)
    print('  ZIP → MSG Converter')
    print('=' * 52)

    if not zip_files:
        print(f'\n[오류] 같은 폴더에 .zip 파일이 없습니다.')
        print(f'  폴더: {base_dir}')
        input('\nEnter 키를 눌러 종료...')
        sys.exit(1)

    print(f'\n  {len(zip_files)}개 파일 발견\n')

    ok = fail = 0
    for zip_path in zip_files:
        base     = os.path.splitext(zip_path)[0]
        msg_path = base + '.msg'
        fname    = os.path.basename(zip_path)
        try:
            out = build_zip_msg(zip_path)
            with open(msg_path, 'wb') as f:
                f.write(out)
            print(f'  ✓  {fname}  →  {os.path.basename(msg_path)}')
            ok += 1
        except Exception as e:
            print(f'  ✗  {fname}  ({e})')
            fail += 1

    print(f'\n  완료: 성공 {ok}개 / 실패 {fail}개')
    input('\nEnter 키를 눌러 종료...')


if __name__ == '__main__':
    main()
